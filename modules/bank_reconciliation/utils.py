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

# Reglas detalladas para clasificación canónica de gastos bancarios.
# Cada entrada es (keywords_a_buscar_en_el_texto, nombre_canónico).
# Se evalúan en orden; la primera que coincide gana.
# Las keywords se verifican en mayúsculas (case-insensitive).
CAT_RULES: tuple[tuple[tuple[str, ...], str], ...] = (
    (("SIRCREB",),                    "Recaudacion SIRCREB CM - Condicion S"),
    (("LEY 25", "DEBITO"),            "Impuesto Ley 25.413 Alic Gral s/Debitos"),
    (("LEY 25", "CREDITO"),           "Impuesto Ley 25.413 Ali Gral s/Creditos"),
    (("LEY 25", "ALIC"),              "Impuesto Ley 25.413 Alic Gral s/Debitos"),
    (("LEY 25", "ALI "),              "Impuesto Ley 25.413 Ali Gral s/Creditos"),
    (("DEBITO FISCAL",),              "I.V.A. - Debito Fiscal 21%"),
    (("I.V .A",),                     "I.V.A. - Debito Fiscal 21%"),
    (("I.V.A",),                      "I.V.A. - Debito Fiscal 21%"),
    (("IVA RG",),                     "Percepcion IVA RG 2408 s/Comis-Gastos"),
    (("PERCEP", "IVA"),               "Percepcion IVA RG 2408 s/Comis-Gastos"),
    (("PERCEP", "BRUTOS"),            "Percep. Ingr. Brutos CABA - Condicion P"),
    (("BRUTOS", "CABA"),              "Percep. Ingr. Brutos CABA - Condicion P"),
    (("INGR BRUTOS",),                "Percep. Ingr. Brutos CABA - Condicion P"),
    (("TUCUMAN",),                    "Recaudacion I.B Tucuman - Condicion J"),
    (("I B TUCUMAN",),                "Recaudacion I.B Tucuman - Condicion J"),
    (("RECAUDACION", "I.B"),          "Recaudacion I.B"),
    (("ECHQ", "CON FILIAL"),          "CHEQUE-Com Acred Camara con Filial Bco"),
    (("ECHQ", "SIN FILIAL"),          "ECHQ-Comis acred Camara sin Filial Bco"),
    (("COMISION POR TRANSFERENCIA",), "Comision por Transferencia"),
    (("COMISION CHEQUE",),            "Comision E-CHEQ pagado por Clearing"),
    (("MANTENIMIENTO",),              "Com. mantenimiento cuenta"),
    (("EXTERIOR", "GIRO"),            "Com. Exterior - Giros y Transferencias"),
    (("EXTERIOR", "IMPORT"),          "Com. Exterior-Operacion de Importacion"),
    (("EXTERIOR", "COMISION"),        "Com. Exterior - Comisiones"),
    (("PAGO DE CHEQUE",),             "Pago de Cheque de Camara"),
    (("RECHAZADO",),                  "Comision Cheque Rechazado"),
    (("COMISION", "RECHAZ"),          "Comision Cheque Rechazado"),
)


def parse_importe(texto: str) -> float:
    """Convierte un importe con formato argentino (``1.234,56``) a float."""
    return float(texto.replace(".", "").replace(",", "."))


def categoria_banco_detalle(desc: str | None) -> str | None:
    """Devuelve el nombre canónico del gasto bancario, o ``None`` si es operación.

    Evalúa ``CAT_RULES`` en orden y retorna la primera categoría cuyas keywords
    aparecen todas en el texto (case-insensitive).  Si ninguna regla coincide
    pero la descripción contiene algún término genérico de ``CARGOS``, devuelve
    ``CATEGORIA_GASTO`` como etiqueta de reserva.
    """
    texto = (desc or "").upper()
    for keywords, categoria in CAT_RULES:
        if all(kw.upper() in texto for kw in keywords):
            return categoria
    # Fallback: términos genéricos (ej. "GUV", "Custodia de Titulos")
    if any(cargo.lower() in (desc or "").lower() for cargo in CARGOS):
        return CATEGORIA_GASTO
    return None


def categoria_banco(desc: str | None) -> str:
    """Clasifica un movimiento del banco según su descripción.

    Returns:
        ``CATEGORIA_GASTO`` si la descripción contiene algún concepto de
        cargo bancario; ``CATEGORIA_OPERACION`` en caso contrario.
    """
    return CATEGORIA_GASTO if categoria_banco_detalle(desc) else CATEGORIA_OPERACION


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
    """DataFrame de gastos/impuestos bancarios pendientes de registrar.

    Agrupa por categoría canónica (``CAT_RULES``) cuando hay coincidencia;
    usa la descripción raw del banco como fallback para gastos no clasificados.
    """
    resumen: dict[str, list] = {}
    for j, b in enumerate(banco):
        if banco_match[j]:
            continue
        label = categoria_banco_detalle(b["desc"])
        if label is None:
            continue
        acc = resumen.setdefault(label, [0, 0.0])
        acc[0] += 1
        acc[1] += b["importe"]
    filas = [
        {"Concepto": label, "Cant. movimientos": cant, "Importe total": imp}
        for label, (cant, imp) in sorted(resumen.items(), key=lambda x: -x[1][1])
    ]
    return pd.DataFrame(filas)


def filter_df(df: pd.DataFrame, estados: set[str]) -> pd.DataFrame:
    """Devuelve las filas cuyo ``Estado`` está en el conjunto dado."""
    if df.empty:
        return df
    return df[df["Estado"].isin(estados)].reset_index(drop=True)
