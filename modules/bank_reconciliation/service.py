"""Servicio de conciliación bancaria.

Punto de entrada del módulo. Orquesta parseo, anulación de autocancelatorios y
matching, y devuelve un :class:`ReconciliationResult` de dominio. **No genera
archivos**: la exportación es responsabilidad de componentes independientes
(ver ``exporters.py`` y ``shared/exporters/``).
"""

from __future__ import annotations

import time
from pathlib import Path
from typing import Optional

import pandas as pd

from shared.logging_config import get_logger

from . import engine, utils
from .models import (
    PeriodInfo,
    ReconciliationData,
    ReconciliationResult,
    ReconciliationStats,
)
from .validators import validate_inputs

logger = get_logger(__name__)


def _build_stats(data: ReconciliationData) -> ReconciliationStats:
    """Calcula las estadísticas del proceso (fórmulas idénticas al legacy)."""
    conta, banco = data.conta, data.banco
    conta_match, banco_match = data.conta_match, data.banco_match

    nc_tot = len(conta)
    nc_ok = sum(1 for x in conta_match if x)
    nb_tot = len(banco)
    nb_ok = sum(1 for x in banco_match if x)

    total_debe = sum(c["debe"] for c in conta)
    total_haber = sum(c["haber"] for c in conta)
    total_credito_banco = sum(b["importe"] for b in banco if b["tipo"] == "C")
    total_debito_banco = sum(b["importe"] for b in banco if b["tipo"] == "D")
    importe_conciliado = sum(conta[i]["importe"] for i in range(nc_tot) if conta_match[i])

    banco_pend_oper = sum(
        1 for j in range(nb_tot)
        if not banco_match[j]
        and utils.categoria_banco(banco[j]["desc"]) == utils.CATEGORIA_OPERACION
    )
    banco_pend_gastos = (nb_tot - nb_ok) - banco_pend_oper

    # Gastos a registrar (misma agregación que el legacy).
    gastos_df = utils.build_gastos_df(banco, banco_match)
    gastos_cant = int(gastos_df["Cant. movimientos"].sum()) if not gastos_df.empty else 0
    gastos_imp = float(gastos_df["Importe total"].sum()) if not gastos_df.empty else 0.0

    saldo_banco, saldo_calc, dif = _saldo_reconciliation(data)

    return ReconciliationStats(
        movimientos_sap=nc_tot,
        sap_conciliados=nc_ok,
        sap_pendientes=nc_tot - nc_ok,
        total_debe=total_debe,
        total_haber=total_haber,
        movimientos_banco=nb_tot,
        banco_conciliados=nb_ok,
        banco_pendientes=nb_tot - nb_ok,
        banco_pend_operaciones=banco_pend_oper,
        banco_pend_gastos=banco_pend_gastos,
        total_debito_banco=total_debito_banco,
        total_credito_banco=total_credito_banco,
        conciliados_por_suma=len(data.grupos),
        gastos_cantidad=gastos_cant,
        gastos_importe=gastos_imp,
        importe_conciliado=importe_conciliado,
        saldo_banco=saldo_banco,
        saldo_contable=data.saldo_contable,
        saldo_calculado=saldo_calc,
        diferencia=dif,
    )


def _saldo_reconciliation(data: ReconciliationData) -> tuple[float, float, float]:
    """Conciliación de saldos (idéntica al legacy generar_resumen_saldos)."""
    conta, banco = data.conta, data.banco
    grupo_banco = {j for js in data.grupos.values() for j in js}

    conta_pend = [c for i, c in enumerate(conta)
                  if not c.get("anulado") and not data.conta_match[i]]
    banco_pend = [b for j, b in enumerate(banco)
                  if not b.get("anulado") and not data.banco_match[j] and j not in grupo_banco]

    t_dep = sum(c["importe"] for c in conta_pend if c["tipo_bco"] == "C")
    t_cheq = sum(c["importe"] for c in conta_pend if c["tipo_bco"] == "D")
    t_cred = sum(b["importe"] for b in banco_pend if b["tipo"] == "C")
    t_deb = sum(b["importe"] for b in banco_pend if b["tipo"] == "D")

    saldo_banco = data.saldo_banco or 0.0
    saldo_calc = saldo_banco + t_dep - t_cheq - t_cred + t_deb
    dif = saldo_calc - data.saldo_contable
    return saldo_banco, saldo_calc, dif


def _build_dataframes(data: ReconciliationData) -> dict[str, pd.DataFrame]:
    """Arma los DataFrames de consumo (UI y otros exportadores)."""
    conta_df = utils.build_contabilidad_df(data.conta, data.conta_match)
    banco_df = utils.build_extracto_df(data.banco, data.banco_match)
    conciliados = {"CONCILIADO", "CONCILIADO POR SUMA"}
    pendientes = {"NO CONCILIADO"}
    return {
        "contabilidad": conta_df,
        "extracto": banco_df,
        "contabilidad_conciliados": utils.filter_df(conta_df, conciliados),
        "contabilidad_pendientes": utils.filter_df(conta_df, pendientes),
        "extracto_conciliados": utils.filter_df(banco_df, conciliados),
        "extracto_pendientes": utils.filter_df(banco_df, pendientes),
        "gastos": utils.build_gastos_df(data.banco, data.banco_match),
    }


def reconcile(
    bank_pdf: Path,
    sap_xlsx: Path,
    saldo_contable: float,
    periodo: str,
    saldo_banco: Optional[float] = None,
) -> ReconciliationResult:
    """Ejecuta la conciliación bancaria y devuelve el resultado de dominio.

    Args:
        bank_pdf: Extracto bancario en PDF.
        sap_xlsx: Libro contable exportado de SAP (Excel).
        saldo_contable: Saldo contable de cierre (obligatorio).
        periodo: Período en formato ``YYYY-MM``.
        saldo_banco: Saldo bancario de cierre. Si es ``None`` se autodetecta
            del PDF.

    Returns:
        ReconciliationResult con estadísticas, DataFrames, payload crudo para
        exportadores y las advertencias/errores del proceso.

    Raises:
        ValidationError: Si las entradas no son válidas.
    """
    started = time.perf_counter()
    bank_pdf, sap_xlsx = Path(bank_pdf), Path(sap_xlsx)
    warnings: list[str] = []

    logger.info("=== Inicio conciliación | período=%s ===", periodo)
    logger.info("Archivos recibidos | banco=%s | sap=%s", bank_pdf.name, sap_xlsx.name)

    validate_inputs(bank_pdf, sap_xlsx, saldo_contable, periodo)
    logger.info("Validación OK | saldo_contable=%.2f", float(saldo_contable))

    period = PeriodInfo.from_string(periodo)

    banco = engine.parse_extracto(bank_pdf, period.stop_marker)
    conta = engine.parse_contabilidad(sap_xlsx)
    logger.info("Parseo OK | banco=%d mov | sap=%d mov", len(banco), len(conta))

    if saldo_banco is None:
        saldo_banco = engine.saldo_final_pdf(bank_pdf, period.stop_marker)
        if saldo_banco is None:
            msg = (f"No se pudo leer el saldo bancario del PDF "
                   f"(marcador '{period.stop_marker}').")
            warnings.append(msg)
            logger.warning(msg)

    conta_anul = engine.marcar_anulados(
        conta, lambda r: "IN" if r["debe"] > 0 else "OUT"
    )
    banco_anul = engine.marcar_anulados(
        banco, lambda r: "IN" if r["tipo"] == "C" else "OUT"
    )
    logger.info("Anulados | conta=%d pares | banco=%d pares",
                len(conta_anul), len(banco_anul))

    conta_match, banco_match, _ = engine.conciliar(conta, banco)
    grupos = engine.conciliar_grupos(conta, banco, conta_match, banco_match)
    logger.info("Matching OK | grupos N:1=%d", len(grupos))

    data = ReconciliationData(
        conta=conta, banco=banco,
        conta_match=conta_match, banco_match=banco_match,
        grupos=grupos, conta_anul=conta_anul, banco_anul=banco_anul,
        period=period,
        saldo_banco=saldo_banco,
        saldo_contable=float(saldo_contable),
    )
    stats = _build_stats(data)
    dataframes = _build_dataframes(data)

    elapsed = time.perf_counter() - started
    logger.info(
        "Conciliación finalizada en %.2fs | conciliados=%d/%d | diferencia=%.2f",
        elapsed, stats.sap_conciliados, stats.movimientos_sap, stats.diferencia,
    )

    return ReconciliationResult(
        stats=stats, data=data, dataframes=dataframes,
        warnings=warnings, errors=[],
    )
