"""Pruebas del motor de conciliación (matching y anulados).

Usan registros sintéticos en memoria (sin PDF/Excel) para validar las reglas de
negocio y protegerlas ante futuras modificaciones.
"""

from datetime import date

from modules.bank_reconciliation import engine


def _conta(fecha, debe=0.0, haber=0.0, clase="AS", numero=1, sujeto="X", desc=""):
    """Crea un registro contable sintético."""
    return {
        "fecha": fecha, "clase": clase, "numero": numero,
        "debe": debe, "haber": haber, "cuit": None, "sujeto": sujeto,
        "desc": desc, "subdiario": None,
        "tipo_bco": "C" if debe > 0 else "D",
        "importe": debe if debe > 0 else haber,
    }


def _banco(fecha, tipo, importe, desc="Transferencia", combte=""):
    """Crea un registro bancario sintético."""
    return {"fecha": fecha, "combte": combte, "desc": desc,
            "tipo": tipo, "importe": importe}


def test_conciliacion_uno_a_uno():
    """Un débito contable (crédito bancario) matchea con su crédito en banco."""
    d = date(2025, 7, 10)
    conta = [_conta(d, debe=100.0)]           # tipo_bco = "C"
    banco = [_banco(d, "C", 100.0)]
    engine.marcar_anulados(conta, lambda r: "IN" if r["debe"] > 0 else "OUT")
    engine.marcar_anulados(banco, lambda r: "IN" if r["tipo"] == "C" else "OUT")

    conta_match, banco_match, pares = engine.conciliar(conta, banco)

    assert conta_match == [True]
    assert banco_match == [True]
    assert pares == [(0, 0)]


def test_lado_invertido_no_matchea_mismo_lado():
    """Un débito contable no matchea contra un débito bancario (lado invertido)."""
    d = date(2025, 7, 10)
    conta = [_conta(d, debe=100.0)]           # espera crédito bancario
    banco = [_banco(d, "D", 100.0)]           # débito → no corresponde
    engine.marcar_anulados(conta, lambda r: "IN" if r["debe"] > 0 else "OUT")
    engine.marcar_anulados(banco, lambda r: "IN" if r["tipo"] == "C" else "OUT")

    conta_match, banco_match, _ = engine.conciliar(conta, banco)

    assert conta_match == [False]
    assert banco_match == [False]


def test_desempate_por_fecha_mas_cercana():
    """Ante dos candidatos del mismo importe, gana el de fecha más cercana."""
    conta = [_conta(date(2025, 7, 10), debe=100.0)]
    banco = [
        _banco(date(2025, 7, 20), "C", 100.0),   # lejos
        _banco(date(2025, 7, 11), "C", 100.0),   # cerca → debe ganar
    ]
    engine.marcar_anulados(conta, lambda r: "IN" if r["debe"] > 0 else "OUT")
    engine.marcar_anulados(banco, lambda r: "IN" if r["tipo"] == "C" else "OUT")

    _, _, pares = engine.conciliar(conta, banco)

    assert pares == [(0, 1)]


def test_conciliacion_agrupada_por_suma():
    """Un movimiento contable matchea con la suma de varios del banco (N:1)."""
    conta = [_conta(date(2025, 7, 10), debe=300.0)]
    banco = [
        _banco(date(2025, 7, 10), "C", 100.0),
        _banco(date(2025, 7, 11), "C", 100.0),
        _banco(date(2025, 7, 12), "C", 100.0),
    ]
    engine.marcar_anulados(conta, lambda r: "IN" if r["debe"] > 0 else "OUT")
    engine.marcar_anulados(banco, lambda r: "IN" if r["tipo"] == "C" else "OUT")
    conta_match, banco_match, _ = engine.conciliar(conta, banco)

    grupos = engine.conciliar_grupos(conta, banco, conta_match, banco_match)

    assert conta_match[0] == "GRUPO"
    assert sorted(grupos[0]) == [0, 1, 2]
    assert all(banco_match)


def test_marcar_anulados_netea_pares_opuestos():
    """Un débito y un crédito de igual importe y fecha se anulan entre sí."""
    d = date(2025, 7, 15)
    conta = [_conta(d, debe=500.0), _conta(d, haber=500.0)]

    pares = engine.marcar_anulados(conta, lambda r: "IN" if r["debe"] > 0 else "OUT")

    assert len(pares) == 1
    assert conta[0]["anulado"] is True
    assert conta[1]["anulado"] is True


def test_anulados_se_excluyen_del_matching():
    """Los movimientos anulados no participan de la conciliación."""
    d = date(2025, 7, 15)
    conta = [_conta(d, debe=500.0), _conta(d, haber=500.0)]
    banco = [_banco(d, "C", 500.0)]
    engine.marcar_anulados(conta, lambda r: "IN" if r["debe"] > 0 else "OUT")
    engine.marcar_anulados(banco, lambda r: "IN" if r["tipo"] == "C" else "OUT")

    conta_match, banco_match, pares = engine.conciliar(conta, banco)

    assert pares == []
    assert banco_match == [False]
