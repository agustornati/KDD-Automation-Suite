"""Validación de las entradas del motor de conciliación."""

from __future__ import annotations

from pathlib import Path

from config import settings

from .models import PeriodInfo


class ValidationError(ValueError):
    """Se lanza cuando las entradas de la conciliación no son válidas.

    Attributes:
        errors: Lista de mensajes de error legibles para el usuario.
    """

    def __init__(self, errors: list[str]) -> None:
        self.errors = errors
        super().__init__("; ".join(errors))


def _validate_file(path: Path, extensiones: tuple[str, ...], etiqueta: str) -> list[str]:
    """Valida existencia y extensión de un archivo de entrada."""
    errores: list[str] = []
    if not path.exists():
        errores.append(f"{etiqueta}: el archivo no existe ({path}).")
        return errores
    if not path.is_file():
        errores.append(f"{etiqueta}: la ruta no es un archivo ({path}).")
        return errores
    if path.suffix.lower() not in extensiones:
        esperadas = ", ".join(extensiones)
        errores.append(
            f"{etiqueta}: extensión no soportada '{path.suffix}'. "
            f"Se esperaba: {esperadas}."
        )
    return errores


def validate_inputs(
    bank_pdf: Path,
    sap_xlsx: Path,
    saldo_contable: float,
    periodo: str,
) -> None:
    """Valida todas las entradas de la conciliación.

    Args:
        bank_pdf: Ruta al extracto bancario (PDF).
        sap_xlsx: Ruta al libro contable exportado de SAP (Excel).
        saldo_contable: Saldo contable de cierre (obligatorio).
        periodo: Período en formato ``YYYY-MM``.

    Raises:
        ValidationError: Si alguna entrada no es válida. Reúne todos los
            errores encontrados para reportarlos juntos.
    """
    errores: list[str] = []

    errores += _validate_file(
        bank_pdf, settings.BANK_STATEMENT_EXTENSIONS, "Extracto bancario"
    )
    errores += _validate_file(
        sap_xlsx, settings.SAP_LEDGER_EXTENSIONS, "Libro SAP"
    )

    if saldo_contable is None:
        errores.append("Saldo contable: es obligatorio.")
    else:
        try:
            float(saldo_contable)
        except (TypeError, ValueError):
            errores.append("Saldo contable: debe ser un número.")

    try:
        PeriodInfo.from_string(periodo)
    except ValueError as exc:
        errores.append(f"Período: {exc}")

    if errores:
        raise ValidationError(errores)
