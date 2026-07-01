"""KDD Automation Suite — interfaz web (Streamlit).

Punto de entrada de la plataforma. Esta capa solo se ocupa de la interacción con
el usuario: captura las entradas, delega en el módulo de negocio
(``modules.bank_reconciliation``) y presenta los resultados. **No contiene
lógica de negocio.**
"""

from __future__ import annotations

from pathlib import Path

import streamlit as st

from config import paths, settings
from modules.bank_reconciliation import (
    ExcelExporter,
    ReconciliationResult,
    ValidationError,
    reconcile,
)
from shared.exporters import ZipExporter
from shared.logging_config import get_logger

logger = get_logger(__name__)


def _save_upload(uploaded_file, destino: Path) -> Path:
    """Guarda un archivo subido en disco y devuelve su ruta."""
    destino.parent.mkdir(parents=True, exist_ok=True)
    destino.write_bytes(uploaded_file.getbuffer())
    return destino


def _run_reconciliation(
    bank_file, sap_file, saldo_contable: float, periodo: str
) -> tuple[ReconciliationResult, Path]:
    """Ejecuta la conciliación y arma el ZIP de resultados.

    Orquesta el módulo de negocio y los exportadores; devuelve el resultado de
    dominio y la ruta del ZIP con los tres Excel.
    """
    bank_path = _save_upload(bank_file, paths.UPLOADS_DIR / bank_file.name)
    sap_path = _save_upload(sap_file, paths.UPLOADS_DIR / sap_file.name)

    result = reconcile(bank_path, sap_path, saldo_contable, periodo)

    out_dir = paths.OUTPUTS_DIR / result.period.mmyyyy
    excels = ExcelExporter().export(result, out_dir)
    zip_path = ZipExporter().bundle(
        excels, out_dir / f"Conciliacion_{result.period.mmyyyy}.zip"
    )
    return result, zip_path


def _render_header() -> None:
    st.title(settings.APP_NAME)
    st.caption(settings.APP_TAGLINE)
    st.divider()


def _render_metrics(result: ReconciliationResult) -> None:
    """Muestra las estadísticas del proceso."""
    s = result.stats
    st.subheader("Resultado")
    fila1 = st.columns(3)
    fila1[0].metric("Movimientos Banco", s.movimientos_banco)
    fila1[1].metric("Movimientos SAP", s.movimientos_sap)
    fila1[2].metric("Conciliados", s.sap_conciliados)
    fila2 = st.columns(3)
    fila2[0].metric("Pendientes", s.sap_pendientes)
    fila2[1].metric("Gastos", s.gastos_cantidad)
    fila2[2].metric("Diferencia", f"{s.diferencia:,.2f}")

    if result.warnings:
        for w in result.warnings:
            st.warning(w)


def main() -> None:
    st.set_page_config(page_title=settings.APP_NAME, page_icon="🏦", layout="centered")
    _render_header()

    st.header("Nueva conciliación bancaria")

    periodo = st.text_input(
        "Período", placeholder=settings.PERIOD_FORMAT_HINT,
        help="Mes a conciliar en formato YYYY-MM (ej. 2025-07).",
    )
    bank_file = st.file_uploader("Extracto Bancario (PDF)", type=["pdf"])
    sap_file = st.file_uploader("Libro SAP (Excel)", type=["xlsx", "xlsm"])
    saldo_contable = st.number_input(
        "Saldo Contable", value=None, step=0.01, format="%.2f",
        help="Saldo contable de cierre del libro banco (obligatorio).",
    )

    if st.button("Conciliar", type="primary"):
        if not (periodo and bank_file and sap_file and saldo_contable is not None):
            st.error("Completá el período, ambos archivos y el saldo contable.")
        else:
            try:
                with st.spinner("Conciliando..."):
                    result, zip_path = _run_reconciliation(
                        bank_file, sap_file, float(saldo_contable), periodo
                    )
                st.session_state["result"] = result
                st.session_state["zip_path"] = str(zip_path)
                st.success("Conciliación completada.")
            except ValidationError as exc:
                st.error("Entradas inválidas:")
                for err in exc.errors:
                    st.error(f"• {err}")
            except Exception as exc:  # noqa: BLE001 — feedback al usuario
                logger.exception("Error en la conciliación")
                st.error(f"Ocurrió un error al conciliar: {exc}")

    # Resultado persistido (sobrevive a los reruns de Streamlit).
    if "result" in st.session_state:
        st.divider()
        _render_metrics(st.session_state["result"])
        zip_path = Path(st.session_state["zip_path"])
        if zip_path.exists():
            st.download_button(
                "Descargar resultados",
                data=zip_path.read_bytes(),
                file_name=zip_path.name,
                mime="application/zip",
            )


if __name__ == "__main__":
    main()
