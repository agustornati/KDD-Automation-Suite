"""Modelos de dominio del módulo de conciliación bancaria.

Definen las estructuras que produce el motor. El motor **no** genera archivos:
devuelve un :class:`ReconciliationResult` con toda la información necesaria para
que un exportador (Excel, ZIP, PDF, etc.) construya la salida que corresponda.
"""

from __future__ import annotations

import calendar
from dataclasses import dataclass, field
from datetime import date
from typing import Optional, TypedDict

import pandas as pd

# Nombres de meses en español (índice 1 = enero), usados en títulos de reportes.
MESES_ES: tuple[str, ...] = (
    "", "ENERO", "FEBRERO", "MARZO", "ABRIL", "MAYO", "JUNIO",
    "JULIO", "AGOSTO", "SEPTIEMBRE", "OCTUBRE", "NOVIEMBRE", "DICIEMBRE",
)


# ---------------------------------------------------------------------------
# Registros crudos (misma forma que usa el motor legacy, con type hints)
# ---------------------------------------------------------------------------
class ContaRecord(TypedDict):
    """Movimiento del libro contable (export de SAP)."""

    fecha: Optional[date]
    clase: object
    numero: object
    debe: float
    haber: float
    cuit: object
    sujeto: object
    desc: object
    subdiario: object
    tipo_bco: str          # "C" (crédito) o "D" (débito) — lado bancario esperado
    importe: float
    anulado: bool


class BancoRecord(TypedDict):
    """Movimiento del extracto bancario (PDF)."""

    fecha: Optional[date]
    combte: str
    desc: str
    tipo: str              # "C" (crédito) o "D" (débito)
    importe: float
    anulado: bool


# ---------------------------------------------------------------------------
# Período
# ---------------------------------------------------------------------------
@dataclass(frozen=True)
class PeriodInfo:
    """Datos derivados del período (``YYYY-MM``).

    De acá salen el marcador de corte del PDF, la etiqueta de saldo y los
    nombres de los archivos de salida, sin nada hardcodeado por mes.
    """

    year: int
    month: int

    @classmethod
    def from_string(cls, periodo: str) -> "PeriodInfo":
        """Construye el período a partir de un texto ``YYYY-MM``."""
        try:
            year_str, month_str = periodo.split("-")
            return cls(int(year_str), int(month_str))
        except (ValueError, AttributeError) as exc:
            raise ValueError(
                "El período debe tener el formato YYYY-MM (ej. 2025-07)"
            ) from exc

    @property
    def last_day(self) -> int:
        return calendar.monthrange(self.year, self.month)[1]

    @property
    def mmyyyy(self) -> str:
        return "%02d-%04d" % (self.month, self.year)

    @property
    def month_name(self) -> str:
        return MESES_ES[self.month]

    @property
    def stop_marker(self) -> str:
        """Línea del PDF que marca el fin del extracto (ej. ``SALDO AL 31/07``)."""
        return "SALDO AL %02d/%02d" % (self.last_day, self.month)

    @property
    def saldo_label(self) -> str:
        return "al %02d/%02d/%04d" % (self.last_day, self.month, self.year)


# ---------------------------------------------------------------------------
# Estadísticas
# ---------------------------------------------------------------------------
@dataclass
class ReconciliationStats:
    """Resumen numérico del proceso de conciliación."""

    # Contabilidad (SAP)
    movimientos_sap: int
    sap_conciliados: int
    sap_pendientes: int
    total_debe: float
    total_haber: float

    # Extracto bancario
    movimientos_banco: int
    banco_conciliados: int
    banco_pendientes: int
    banco_pend_operaciones: int
    banco_pend_gastos: int
    total_debito_banco: float
    total_credito_banco: float

    # Matching agrupado y gastos
    conciliados_por_suma: int
    gastos_cantidad: int
    gastos_importe: float
    importe_conciliado: float

    # Conciliación de saldos
    saldo_banco: float
    saldo_contable: float
    saldo_calculado: float
    diferencia: float


# ---------------------------------------------------------------------------
# Payload crudo para los exportadores
# ---------------------------------------------------------------------------
@dataclass
class ReconciliationData:
    """Registros crudos y estado del matching.

    Es lo que necesita un exportador para reconstruir cualquier salida (los
    Excel del legacy, un PDF futuro, etc.) sin volver a ejecutar el motor.
    """

    conta: list[ContaRecord]
    banco: list[BancoRecord]
    conta_match: list          # elementos: False | True | "GRUPO"
    banco_match: list[bool]
    grupos: dict[int, list[int]]
    conta_anul: list[tuple[int, int]]
    banco_anul: list[tuple[int, int]]
    period: PeriodInfo
    saldo_banco: Optional[float]   # None si no se pudo autodetectar del PDF
    saldo_contable: float


# ---------------------------------------------------------------------------
# Resultado de dominio
# ---------------------------------------------------------------------------
@dataclass
class ReconciliationResult:
    """Objeto de dominio devuelto por el motor.

    Contiene estadísticas, DataFrames listos para consumir, el payload crudo
    para los exportadores, y las advertencias/errores del proceso. **No**
    referencia ningún archivo generado: la exportación es responsabilidad de
    componentes independientes.
    """

    stats: ReconciliationStats
    data: ReconciliationData
    dataframes: dict[str, pd.DataFrame]
    warnings: list[str] = field(default_factory=list)
    errors: list[str] = field(default_factory=list)

    @property
    def period(self) -> PeriodInfo:
        return self.data.period

    @property
    def matched(self) -> pd.DataFrame:
        """Movimientos contables conciliados."""
        return self.dataframes["contabilidad_conciliados"]

    @property
    def pending(self) -> pd.DataFrame:
        """Movimientos contables pendientes."""
        return self.dataframes["contabilidad_pendientes"]

    @property
    def is_balanced(self) -> bool:
        """True si la diferencia de saldos es cero (a dos decimales)."""
        return round(self.stats.diferencia, 2) == 0.0
