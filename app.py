"""NAT Automation Suite — interfaz web (Streamlit).

Punto de entrada de la plataforma. Esta capa solo se ocupa de la experiencia de
usuario: captura las entradas, delega en el módulo de negocio
(``modules.bank_reconciliation``) y presenta los resultados de forma clara para
usuarios no técnicos. **No contiene lógica de negocio** ni modifica el motor.
"""

from __future__ import annotations

import time
from datetime import datetime
from pathlib import Path

import streamlit as st

from config import paths, settings
from modules.bank_reconciliation import (
    ExcelExporter,
    ReconciliationResult,
    ValidationError,
    reconcile,
)
from modules.bank_reconciliation.engine import is_scanned_pdf
from shared.exporters import ZipExporter
from shared.formatting import format_ars, human_size
from shared.logging_config import get_logger

logger = get_logger(__name__)

# CSS mínimo: cajas de carga más grandes, tarjetas de métricas y espaciado.
_CSS = """
<style>
[data-testid="stFileUploaderDropzone"] {
    border: 2px dashed #b9c0d4;
    border-radius: 10px;
    padding: 26px;
    background-color: #fafbfe;
}
[data-testid="stMetric"] {
    background-color: #ffffff;
    border: 1px solid #e6e8ef;
    border-radius: 10px;
    padding: 14px 16px;
}
div.block-container { padding-top: 2.2rem; }
</style>
"""


# ---------------------------------------------------------------------------
# Helpers de UI
# ---------------------------------------------------------------------------
def _save_upload(uploaded_file, destino: Path) -> Path:
    """Guarda un archivo subido en disco y devuelve su ruta."""
    destino.parent.mkdir(parents=True, exist_ok=True)
    destino.write_bytes(uploaded_file.getbuffer())
    return destino


def _friendly_error(exc: Exception) -> str:
    """Traduce una excepción técnica a un mensaje claro para el usuario.

    El detalle completo se registra en ``logs/``; acá solo devolvemos un texto
    amigable, sin tracebacks ni jerga de Python.
    """
    origen = (type(exc).__module__ or "").lower()
    nombre = type(exc).__name__.lower()
    if "pdf" in origen or "pdf" in nombre:
        return ("No se pudo leer el PDF seleccionado. "
                "Verificá que sea un extracto bancario válido.")
    if "openpyxl" in origen or "xl" in nombre or "excel" in nombre:
        return ("El archivo contable no tiene el formato esperado. "
                "Verificá que sea el Excel del libro contable.")
    return ("Ocurrió un problema al procesar la conciliación. "
            "Revisá los archivos e intentá nuevamente.")


def _missing_fields(periodo: str, bank_file, sap_file, saldo) -> list[str]:
    """Devuelve la lista de campos obligatorios que faltan completar."""
    faltantes: list[str] = []
    if not periodo:
        faltantes.append("el **período** (formato AAAA-MM)")
    if bank_file is None:
        faltantes.append("el **extracto bancario** (PDF)")
    if sap_file is None:
        faltantes.append("los **libros contables** (Excel)")
    if saldo is None:
        faltantes.append("el **saldo contable**")
    return faltantes


def _process(bank_file, sap_file, saldo: float, periodo: str, saldo_banco: float | None = None):
    """Ejecuta la conciliación mostrando el progreso por etapas.

    Orquesta el módulo de negocio y los exportadores. Devuelve el resultado de
    dominio, la ruta del ZIP, la lista de Excel y el tiempo total.
    """
    started = time.perf_counter()
    with st.status("Procesando conciliación...", expanded=True) as status:
        st.write("🔍 Validando archivos...")
        bank_path = _save_upload(bank_file, paths.UPLOADS_DIR / bank_file.name)
        sap_path = _save_upload(sap_file, paths.UPLOADS_DIR / sap_file.name)

        if is_scanned_pdf(bank_path):
            st.write("🔎 PDF escaneado detectado — aplicando OCR (puede tardar 5–10 min)...")
            if saldo_banco is None:
                raise ValidationError(["Para PDFs escaneados (ej. Credicoop) el "
                                        "saldo bancario al cierre es obligatorio."])
        else:
            st.write("📄 Leyendo extracto y procesando contabilidad...")
        result = reconcile(bank_path, sap_path, saldo, periodo, saldo_banco)

        st.write("📦 Generando resultados (Excel + ZIP)...")
        out_dir = paths.OUTPUTS_DIR / result.period.mmyyyy
        excels = ExcelExporter().export(result, out_dir)
        zip_path = ZipExporter().bundle(
            excels, out_dir / f"Conciliacion_{result.period.mmyyyy}.zip"
        )

        st.write("🏁 Finalizando...")
        elapsed = time.perf_counter() - started
        status.update(
            label=f"Conciliación completada en {elapsed:.1f} s",
            state="complete", expanded=False,
        )
    return result, zip_path, excels, elapsed


def _build_log(result: ReconciliationResult, bank_name: str, sap_name: str,
               n_files: int, zip_size: int, elapsed: float) -> list[str]:
    """Arma los mensajes importantes del proceso para el log de ejecución."""
    s = result.stats
    msgs = [
        f"Archivos recibidos: {bank_name} · {sap_name}",
        f"Banco: {s.movimientos_banco} movimientos · Contabilidad: {s.movimientos_sap} movimientos",
        f"Conciliados: {s.sap_conciliados} · Pendientes: {s.sap_pendientes} · "
        f"Por grupo: {s.conciliados_por_suma}",
        f"Diferencia de saldos: {format_ars(s.diferencia)}",
    ]
    for w in result.warnings:
        msgs.append(f"Aviso: {w}")
    msgs.append(f"Resultados: {n_files} archivos ({human_size(zip_size)}) en {elapsed:.1f} s")
    return msgs


# ---------------------------------------------------------------------------
# Secciones
# ---------------------------------------------------------------------------
def _render_header() -> None:
    """Encabezado: logo (opcional), nombre, versión y última ejecución."""
    logo = paths.ASSETS_DIR / "logo.png"
    if logo.exists():
        col_logo, col_title = st.columns([1, 6])
        col_logo.image(str(logo), width=72)
        container = col_title
    else:
        container = st

    container.title(settings.APP_NAME)
    container.caption(f"{settings.APP_TAGLINE}  ·  v{settings.APP_VERSION}")

    last_run = st.session_state.get("last_run_at")
    if last_run:
        st.caption(f"🕒 Última ejecución: {last_run:%d/%m/%Y %H:%M:%S}")
    st.divider()


def _render_file_summary(uploaded_file) -> None:
    """Muestra nombre, tamaño y estado de un archivo cargado."""
    if uploaded_file is not None:
        st.success(f"✔ {uploaded_file.name}  ·  {human_size(uploaded_file.size)}")


def _render_new_reconciliation():
    """Sección de entrada. Devuelve (periodo, bank_file, sap_file, saldo)."""
    st.subheader("Nueva conciliación bancaria")

    periodo = st.text_input(
        "Período", key="period", placeholder=settings.PERIOD_FORMAT_HINT,
        help="Mes a conciliar en formato AAAA-MM (ej. 2025-07).",
    )

    st.markdown("**Extracto bancario (PDF)** — arrastre el archivo aquí o haga clic para seleccionarlo")
    bank_file = st.file_uploader(
        "Extracto bancario (PDF)", type=["pdf"],
        label_visibility="collapsed", key="bank_upload",
    )
    _render_file_summary(bank_file)

    st.markdown("**Libros Contables (Excel)** — arrastre el archivo aquí o haga clic para seleccionarlo")
    sap_file = st.file_uploader(
        "Libros Contables (Excel)", type=["xlsx", "xlsm"],
        label_visibility="collapsed", key="sap_upload",
    )
    _render_file_summary(sap_file)

    saldo = st.number_input(
        "Saldo contable", key="saldo", value=None, step=0.01, format="%.2f",
        help="Saldo contable de cierre del libro banco (obligatorio).",
    )
    saldo_banco = st.number_input(
        "Saldo bancario al cierre", key="saldo_banco", value=None,
        step=0.01, format="%.2f",
        help="Saldo del extracto bancario al último día del período. "
             "Obligatorio para PDFs escaneados (ej. Credicoop): el OCR no puede "
             "leer el saldo correctamente en ese formato.",
    )
    if saldo_banco is None:
        st.caption("⚠️ Obligatorio para extractos Credicoop y otros PDFs escaneados. "
                   "Sin este valor el OCR no puede calcular correctamente los movimientos faltantes.")
    return periodo, bank_file, sap_file, saldo, saldo_banco


def _render_results(result: ReconciliationResult, elapsed: float) -> None:
    """Muestra las estadísticas como tarjetas con iconos y el resumen de éxito."""
    s = result.stats
    st.subheader("Resultado")

    fila1 = st.columns(3)
    fila1[0].metric("🏦 Movimientos Banco", s.movimientos_banco)
    fila1[1].metric("📊 Movimientos Contables", s.movimientos_sap)
    fila1[2].metric("✅ Conciliados", s.sap_conciliados)
    fila2 = st.columns(3)
    fila2[0].metric("⏳ Pendientes", s.sap_pendientes)
    fila2[1].metric("🔗 Conciliaciones por grupo", s.conciliados_por_suma)
    fila2[2].metric("⚖️ Diferencia", format_ars(s.diferencia))

    procesados = s.movimientos_banco + s.movimientos_sap
    st.success(
        f"**La conciliación finalizó correctamente.**\n\n"
        f"- ⏱️ Tiempo de ejecución: **{elapsed:.1f} s**\n"
        f"- 📁 Registros procesados: **{procesados}**\n"
        f"- ⏳ Pendientes: **{s.sap_pendientes}**  ·  "
        f"⚖️ Diferencia de saldos: **{format_ars(s.diferencia)}**"
    )

    for w in result.warnings:
        st.warning(w)


def _render_download(zip_path: Path, n_files: int) -> None:
    """Sección independiente de descarga de resultados."""
    st.subheader("Resultados disponibles")
    if not zip_path.exists():
        return
    size = zip_path.stat().st_size
    st.caption(f"📦 {n_files} archivos incluidos  ·  tamaño total {human_size(size)}")
    st.download_button(
        "⬇️  Descargar ZIP",
        data=zip_path.read_bytes(),
        file_name=zip_path.name,
        mime="application/zip",
        type="primary",
    )


def _render_log(msgs: list[str]) -> None:
    """Log de ejecución con los mensajes importantes del proceso."""
    st.subheader("Log de ejecución")
    with st.container(border=True):
        for m in msgs:
            st.markdown(f"- {m}")


# ---------------------------------------------------------------------------
# Aplicación
# ---------------------------------------------------------------------------
def main() -> None:
    st.set_page_config(page_title=settings.APP_NAME, page_icon="⚡", layout="centered")
    st.markdown(_CSS, unsafe_allow_html=True)

    _render_header()
    periodo, bank_file, sap_file, saldo, saldo_banco = _render_new_reconciliation()

    if st.button("Conciliar", type="primary", use_container_width=True):
        faltantes = _missing_fields(periodo, bank_file, sap_file, saldo)
        if faltantes:
            st.error("Para conciliar, primero completá: " + "; ".join(faltantes) + ".")
        else:
            try:
                result, zip_path, excels, elapsed = _process(
                    bank_file, sap_file, float(saldo), periodo,
                    float(saldo_banco) if saldo_banco is not None else None,
                )
                st.session_state["result"] = result
                st.session_state["zip_path"] = str(zip_path)
                st.session_state["n_files"] = len(excels)
                st.session_state["elapsed"] = elapsed
                st.session_state["last_run_at"] = datetime.now()
                st.session_state["log_msgs"] = _build_log(
                    result, bank_file.name, sap_file.name,
                    len(excels), zip_path.stat().st_size, elapsed,
                )
            except ValidationError as exc:
                st.error("No se pudo iniciar la conciliación:")
                for err in exc.errors:
                    st.error(f"• {err}")
            except Exception as exc:  # noqa: BLE001 — feedback amigable al usuario
                logger.exception("Error en la conciliación")
                st.error(_friendly_error(exc))

    # Resultado persistido (sobrevive a los reruns de Streamlit).
    if "result" in st.session_state:
        st.divider()
        _render_results(st.session_state["result"], st.session_state["elapsed"])
        st.divider()
        _render_download(
            Path(st.session_state["zip_path"]), st.session_state["n_files"]
        )
        st.divider()
        _render_log(st.session_state["log_msgs"])


if __name__ == "__main__":
    main()
