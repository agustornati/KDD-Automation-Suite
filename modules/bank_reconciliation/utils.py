"""Utilidades del motor de conciliación.

Helpers de dominio (parseo de importes, clasificación de movimientos bancarios)
y construcción de DataFrames a partir de los registros crudos. Toda la lógica
proviene del script legacy y se mantiene con idéntico comportamiento.
"""

from __future__ import annotations

import re

import pandas as pd

from .models import BancoRecord, ContaRecord, PeriodInfo

# Expresiones de reconocimiento de importes y fechas en el PDF del extracto.
AMOUNT_RE = re.compile(r"^\d{1,3}(?:\.\d{3})*,\d{2}$")
DATE_RE = re.compile(r"^\d{2}/\d{2}/\d{2}$")

# Conceptos que identifican gastos/impuestos bancarios (no operaciones).
CARGOS: tuple[str, ...] = (
    "Impuesto Ley", "SIRCREB", "Recaudacion I.B", "I.V.A.",
    "Percepcion", "Percep.", "Comision", "Com. Exterior",
    "Com. mantenimiento", "GUV", "Comis acred", "Com Acred",
    "Comision E-CHEQ", "Custodia de Titulos",
)

CATEGORIA_GASTO = "Gasto/Impuesto bancario"
CATEGORIA_OPERACION = "Operacion"


def parse_importe(texto: str) -> float:
    """Convierte un importe con formato argentino (``1.234,56``) a float."""
    return float(texto.replace(".", "").replace(",", "."))


def categoria_banco(desc: str | None) -> str:
    """Clasifica un movimiento del banco según su descripción.

    Returns:
        ``CATEGORIA_GASTO`` si la descripción contiene algún concepto de
        cargo bancario; ``CATEGORIA_OPERACION`` en caso contrario.
    """
    texto = desc or ""
    if any(cargo.lower() in texto.lower() for cargo in CARGOS):
        return CATEGORIA_GASTO
    return CATEGORIA_OPERACION


def _lado_banco(tipo: str) -> str:
    """Traduce el código de lado bancario a etiqueta legible."""
    return "Credito" if tipo == "C" else "Debito"


def _estado(match_value: object) -> str:
    """Traduce el valor de matching a etiqueta de estado."""
    if match_value == "GRUPO":
        return "CONCILIADO POR SUMA"
    return "CONCILIADO" if match_value else "NO CONCILIADO"


# ---------------------------------------------------------------------------
# Construcción de DataFrames para consumo de la UI / otros exportadores
# ---------------------------------------------------------------------------
def build_contabilidad_df(
    conta: list[ContaRecord], conta_match: list
) -> pd.DataFrame:
    """DataFrame de todos los movimientos contables (excluye anulados)."""
    filas = [
        {
            "Fecha": c["fecha"],
            "Clase": c["clase"],
            "Numero": c["numero"],
            "Debe": c["debe"] or None,
            "Haber": c["haber"] or None,
            "Sujeto": c["sujeto"],
            "Lado banco": _lado_banco(c["tipo_bco"]),
            "Estado": _estado(conta_match[i]),
        }
        for i, c in enumerate(conta)
        if not c["anulado"]
    ]
    return pd.DataFrame(filas)


def build_extracto_df(
    banco: list[BancoRecord], banco_match: list[bool]
) -> pd.DataFrame:
    """DataFrame de todos los movimientos bancarios (excluye anulados)."""
    filas = [
        {
            "Fecha": b["fecha"],
            "Combte": b["combte"],
            "Descripcion": b["desc"],
            "Debito": b["importe"] if b["tipo"] == "D" else None,
            "Credito": b["importe"] if b["tipo"] == "C" else None,
            "Estado": _estado(banco_match[j]),
            "Categoria": categoria_banco(b["desc"]),
        }
        for j, b in enumerate(banco)
        if not b["anulado"]
    ]
    return pd.DataFrame(filas)


def build_gastos_df(banco: list[BancoRecord], banco_match: list[bool]) -> pd.DataFrame:
    """DataFrame de gastos/impuestos bancarios pendientes de registrar."""
    resumen: dict[str, list] = {}
    for j, b in enumerate(banco):
        if banco_match[j]:
            continue
        if categoria_banco(b["desc"]) != CATEGORIA_GASTO:
            continue
        acc = resumen.setdefault(b["desc"], [0, 0.0])
        acc[0] += 1
        acc[1] += b["importe"]
    filas = [
        {"Concepto": desc, "Cant. movimientos": cant, "Importe total": imp}
        for desc, (cant, imp) in sorted(resumen.items(), key=lambda x: -x[1][1])
    ]
    return pd.DataFrame(filas)


def filter_df(df: pd.DataFrame, estados: set[str]) -> pd.DataFrame:
    """Devuelve las filas cuyo ``Estado`` está en el conjunto dado."""
    if df.empty:
        return df
    return df[df["Estado"].isin(estados)].reset_index(drop=True)
