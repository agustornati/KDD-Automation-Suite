"""Smoke tests de la interfaz (Streamlit AppTest).

No dependen de archivos de clientes: validan que la app cargue sin errores y que
la validación de campos obligatorios muestre un mensaje claro.
"""

from streamlit.testing.v1 import AppTest

TIMEOUT = 30


def _app() -> AppTest:
    return AppTest.from_file("app.py")


def test_app_carga_sin_errores():
    at = _app().run(timeout=TIMEOUT)
    assert not at.exception
    assert at.title[0].value == "KDD Automation Suite"


def test_conciliar_sin_datos_muestra_validacion():
    at = _app().run(timeout=TIMEOUT)
    at.button[0].click().run(timeout=TIMEOUT)
    assert not at.exception
    assert len(at.error) >= 1
    assert "completá" in at.error[0].value.lower()


def test_formato_numero_argentino():
    from shared.formatting import format_ars, human_size

    assert format_ars(239716.74) == "239.716,74"
    assert format_ars(1234567.8) == "1.234.567,80"
    assert human_size(0) == "0 B"
    assert human_size(2048) == "2.0 KB"
