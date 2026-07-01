"""Pruebas de utilidades de dominio y del período."""

import pytest

from modules.bank_reconciliation.models import PeriodInfo
from modules.bank_reconciliation.utils import (
    CATEGORIA_GASTO,
    CATEGORIA_OPERACION,
    categoria_banco,
    parse_importe,
)


def test_parse_importe_formato_argentino():
    assert parse_importe("1.234.567,89") == 1234567.89
    assert parse_importe("0,50") == 0.50


def test_categoria_banco_detecta_gastos():
    assert categoria_banco("Impuesto Ley 25.413") == CATEGORIA_GASTO
    assert categoria_banco("SIRCREB retencion") == CATEGORIA_GASTO
    assert categoria_banco("Comision mantenimiento") == CATEGORIA_GASTO


def test_categoria_banco_operacion_por_defecto():
    assert categoria_banco("Transferencia recibida") == CATEGORIA_OPERACION
    assert categoria_banco("") == CATEGORIA_OPERACION
    assert categoria_banco(None) == CATEGORIA_OPERACION


def test_period_info_derivaciones():
    p = PeriodInfo.from_string("2025-07")
    assert p.year == 2025
    assert p.month == 7
    assert p.last_day == 31
    assert p.mmyyyy == "07-2025"
    assert p.month_name == "JULIO"
    assert p.stop_marker == "SALDO AL 31/07"
    assert p.saldo_label == "al 31/07/2025"


def test_period_info_formato_invalido():
    with pytest.raises(ValueError):
        PeriodInfo.from_string("2025/07")
