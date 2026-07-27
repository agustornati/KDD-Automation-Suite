"""Motor de conciliación bancaria (parseo + matching).

Contiene la lógica de negocio migrada del script legacy
(``legacy/conciliacion_mvp.py``) **sin cambios de comportamiento**: parseo del
extracto (PDF) y del libro contable (Excel), anulación de autocancelatorios,
matching uno-a-uno y matching agrupado N:1 por suma.

Regla contable: Contabilidad DEBE (entrada) ↔ Banco CRÉDITO ;
                Contabilidad HABER (salida)  ↔ Banco DÉBITO (lado invertido).
"""

from __future__ import annotations

import re
from collections import defaultdict
from datetime import date, datetime
from itertools import combinations
from pathlib import Path
from statistics import median as _statistics_median
from typing import Callable, Optional

import openpyxl
import pdfplumber

from shared.logging_config import get_logger
from .models import BancoRecord, ContaRecord

logger = get_logger(__name__)
from .utils import (
    AMOUNT_RE,
    CATEGORIA_OPERACION,
    DATE_RE,
    categoria_banco,
    parse_importe,
)


# ===========================================================================
# 1. Parseo del EXTRACTO (PDF) por coordenadas
# ===========================================================================

# --- OCR fallback (PDFs escaneados / imagen) --------------------------------

_OCR_ZOOM = 3          # factor de zoom para renderizado antes del OCR
_DEB_X1   = 1281       # borde derecho columna DEBITO  (a 3x = 854*1.5)
_CRED_X1  = 1590       # borde derecho columna CREDITO (a 3x = 1060*1.5)
_DATE_X0  = 300        # borde derecho columna FECHA   (a 3x = 200*1.5)
_COMBTE_X0 = 195       # inicio columna COMBTE         (a 3x = 130*1.5)
_COMBTE_X1 = 345       # fin columna COMBTE            (a 3x = 230*1.5)
_AMT_X0   = 975        # inicio zona de importes       (a 3x = 650*1.5)
_SALDO_X0_WIDE = 1560  # zona SALDO ampliada           (a 3x = 1040*1.5)
_SALDO_RECOVERY_MIN = 5_000.0  # discrepancia mínima ARS para activar recuperación
# SIRCREB: si un débito supera N veces la mediana del mes, es un misread del OCR
# (ej: OCR lee el saldo corriente como importe de la retención → cifras de 9M, 5M)
_SIRCREB_OUTLIER_FACTOR = 100

# Fecha en zona de importes: "DD/MM/YYYY" o "DD/MM/YY" leída como crédito/débito
_DATE_TOKEN_RE = re.compile(r"^\d{2}/\d{2}/\d{2,4}$")

# Patrones en descripciones que indican líneas de encabezado/pie de página a ignorar
_SKIP_LINE_RE = re.compile(
    r"VIENE\s*DE\s*(PAG|PAGINA)|PAGINA\s*ANTERIOR|TOTALES?\s*DEL?\s*DIA",
    re.IGNORECASE,
)


def _filter_sircreb_outliers(movs: list[dict]) -> list[dict]:
    """Descarta movimientos SIRCREB cuyo importe es un misread del OCR.

    El OCR a veces confunde la columna de saldo con la columna de importe y
    genera un débito SIRCREB con el valor del saldo corriente (ej: 9.000.000).
    Si un registro supera _SIRCREB_OUTLIER_FACTOR veces la mediana del mes, se
    elimina y se registra un warning; el algoritmo SALDO-RECOVER lo compensa.
    """
    importes = [m["importe"] for m in movs
                if "SIRCREB" in (m.get("desc") or "").upper() and m["importe"] > 0]
    if len(importes) < 3:
        return movs
    med = _statistics_median(importes)
    umbral = max(med * _SIRCREB_OUTLIER_FACTOR, 1_000_000.0)
    out: list[dict] = []
    for m in movs:
        if ("SIRCREB" in (m.get("desc") or "").upper() and m["importe"] > umbral):
            logger.warning(
                "SIRCREB outlier descartado (OCR misread) | fecha=%s importe=%.2f "
                "mediana=%.2f umbral=%.2f",
                m.get("fecha"), m["importe"], med, umbral,
            )
        else:
            out.append(m)
    return out


def _recover_from_saldo(
    movs_with_y: list[tuple[int, dict]],
    saldo_checks: list[tuple[int, float]],
    fallback_date: Optional[date],
) -> list[dict]:
    """Genera movimientos sintéticos para débitos/créditos que OCR no leyó.

    Cuando la columna SALDO indica un cambio que los movimientos OCR no explican
    (diferencia > _SALDO_RECOVERY_MIN), inserta un movimiento "SALDO-RECOVER"
    con el importe de la diferencia.  Estos quedan visibles en el informe de
    conciliación como partidas pendientes a revisar manualmente.
    """
    if len(saldo_checks) < 2:
        return []

    recovered = []
    for idx in range(len(saldo_checks) - 1):
        y_prev, s_prev = saldo_checks[idx]
        y_curr, s_curr = saldo_checks[idx + 1]

        between = [(y, m) for y, m in movs_with_y if y_prev < y <= y_curr]
        ocr_cred = sum(m["importe"] for _, m in between if m["tipo"] == "C")
        ocr_deb  = sum(m["importe"] for _, m in between if m["tipo"] == "D")

        actual_delta = s_curr - s_prev
        ocr_delta    = ocr_cred - ocr_deb
        discrepancy  = actual_delta - ocr_delta  # negativo → débito faltante

        if abs(discrepancy) < _SALDO_RECOVERY_MIN:
            continue

        fecha = next((m["fecha"] for _, m in between if m.get("fecha")), fallback_date)
        n_between = len(between)
        if discrepancy < 0:
            # Filtrar DEBs pequeños en segmentos sin movimientos OCR:
            # posibles checkpoints de saldo ruidosos (encabezados/resúmenes de página).
            if n_between == 0 and abs(discrepancy) < 15_000_000:
                continue
            recovered.append({
                "fecha": fecha, "combte": "",
                "desc": f"SALDO-RECOVER DEB {abs(discrepancy):,.2f}",
                "tipo": "D", "importe": abs(discrepancy),
            })
        else:
            recovered.append({
                "fecha": fecha, "combte": "",
                "desc": f"SALDO-RECOVER CRED {discrepancy:,.2f}",
                "tipo": "C", "importe": discrepancy,
            })

    return recovered


def _filter_partial_saldo(checks: list[tuple[int, float]]) -> list[tuple[int, float]]:
    """Elimina checkpoints de saldo que parecen lecturas parciales del OCR.

    Síndrome: "98.640.089,06" se lee como "640.089" → checkpoint anómalo rodeado
    por valores 100x mayores.  Criterio: valor < 1% del anterior Y el siguiente
    checkpoint es > 10x el actual.
    """
    if len(checks) < 3:
        return checks
    filtered = [checks[0]]
    for i in range(1, len(checks)):
        curr_y, curr_s = checks[i]
        prev_s = filtered[-1][1]
        is_partial = (
            curr_s < prev_s * 0.01
            and i + 1 < len(checks)
            and checks[i + 1][1] > curr_s * 10
        )
        if not is_partial:
            filtered.append((curr_y, curr_s))
    return filtered


def _get_ocr_reader():
    """Inicializa EasyOCR la primera vez y lo cachea en el módulo."""
    if not hasattr(_get_ocr_reader, "_reader"):
        import easyocr  # importación lazy: solo si el PDF es escaneado
        _get_ocr_reader._reader = easyocr.Reader(
            ["es", "en"], gpu=False, verbose=False
        )
    return _get_ocr_reader._reader


def _normalize_ocr_amount(fragments: list[str]) -> Optional[float]:
    """Reconstruye un importe en formato argentino desde fragmentos OCR.

    Maneja los patrones de OCR sobre Courier vectorizado o escaneado:
      ``'832 .376, 00'``  → 832.376,00  (espacios alrededor de separadores)
      ``'991.300_00'``    → 991.300,00  (guión bajo como separador decimal)
      ``'26.200 6'``      → 26.200,00  (dígito suelto final = artefacto OCR)
      ``'504'``           → 504,00      (decimal truncado por OCR)
    """
    if not fragments:
        return None

    # Token suelto de 1 dígito al final = artefacto OCR (ej: "26.200 6" → 26200)
    if len(fragments) >= 2:
        last = fragments[-1].strip()
        if len(last) == 1 and last.isdigit():
            fragments = fragments[:-1]

    # Si hay múltiples fragmentos que individualmente son importes válidos, significa
    # que dos filas se fusionaron en el mismo bucket y cada una tiene su importe.
    # Tomamos el último (más a la derecha = columna real del importe).
    if len(fragments) >= 2:
        validos = [f.strip() for f in fragments if AMOUNT_RE.match(f.strip())]
        if len(validos) >= 2:
            return parse_importe(validos[-1])
        if len(validos) == 1:
            return parse_importe(validos[0])

    # Unir con espacio, normalizar separadores garbled, luego eliminar espacios
    raw = " ".join(fragments)
    raw = raw.replace("_", ".").replace("|", ",")   # _ = miles garbled, | = coma garbled
    clean = re.sub(r"\s+", "", raw)
    clean = re.sub(r"[^\d.,]", "", clean)           # solo dígitos, puntos, comas

    if not clean:
        return None

    # Intento 1: ya cumple formato argentino exacto N.NNN,NN
    if AMOUNT_RE.match(clean):
        return parse_importe(clean)

    # Intento 2: última coma como separador decimal (maneja "1.234,5" o comas múltiples)
    if "," in clean:
        ridx = clean.rfind(",")
        int_part = clean[:ridx]
        dec_part = clean[ridx + 1:]
        if 1 <= len(dec_part) <= 2 and dec_part.isdigit() and re.match(r"^[\d.]+$", int_part):
            int_str = int_part.replace(".", "")
            if int_str.isdigit() and len(int_str) <= 10:
                return float(int_str + "." + dec_part.ljust(2, "0"))

    # Intento 3: N.NNN...NNN + 2 dígitos finales pegados (decimal truncado al unir)
    m = re.match(r"^(\d{1,3}(?:\.\d{3})+)(\d{2})$", clean)
    if m:
        return float(m.group(1).replace(".", "") + "." + m.group(2))

    # Sin separadores: decimal (",00") truncado por OCR
    digits = re.sub(r"[^\d]", "", clean)
    if digits:
        if len(digits) > 12:  # > 1 billón ARS → concatenación de dos filas
            return None
        # Fecha DDMMYYYY leída como importe (ej: "30/09/2025" → 30092025)
        if len(digits) >= 8:
            try:
                from datetime import date as _d
                _d(int(digits[4:8]), int(digits[2:4]), int(digits[:2]))
                return None
            except (ValueError, TypeError):
                pass
        # > 50M ARS sin separadores (fallback numérico) → casi siempre basura OCR
        # Los importes reales grandes llegan con "." y "," correctos y se parsean antes
        if int(digits) > 50_000_000:
            return None
        return float(digits + ".00")

    return None


def _reconstruct_ocr_date(toks: list) -> Optional[date]:
    """Reconstruye una fecha desde tokens OCR de la columna FECHA."""
    for (_, _, _, text) in toks:
        if DATE_RE.match(text):
            try:
                dd, mm, yy = text.split("/")
                return date(2000 + int(yy), int(mm), int(dd))
            except (ValueError, TypeError):
                pass
    # Intento: tokens parciales pegados (ej. '05/08' + '25')
    joined = "".join(t[3] for t in toks).replace(" ", "")
    m = re.search(r"(\d{2})/(\d{2})/(\d{2})", joined)
    if m:
        try:
            return date(2000 + int(m.group(3)), int(m.group(2)), int(m.group(1)))
        except (ValueError, TypeError):
            pass
    return None


def _ocr_cache_path(pdf_path: Path) -> Path:
    return pdf_path.with_suffix(pdf_path.suffix + ".ocr_cache.json")


def _pdf_hash(pdf_path: Path) -> str:
    """SHA-256 del contenido del PDF. Independiente del timestamp del archivo."""
    import hashlib
    h = hashlib.sha256()
    with open(pdf_path, "rb") as f:
        for chunk in iter(lambda: f.read(65536), b""):
            h.update(chunk)
    return h.hexdigest()


def _load_ocr_cache(
    pdf_path: Path, saldo_final: Optional[float] = None
) -> Optional[tuple[list[BancoRecord], Optional[float]]]:
    import json
    cache_path = _ocr_cache_path(pdf_path)
    if not cache_path.exists():
        return None
    try:
        with open(cache_path, encoding="utf-8") as f:
            data = json.load(f)
        # Clave de identidad: hash del contenido.
        # Compatibilidad: caches viejos usan "mtime"; si coincide, migrar a hash.
        pdf_hash = _pdf_hash(pdf_path)
        if "hash" in data:
            if data["hash"] != pdf_hash:
                return None
        elif "mtime" in data:
            if data["mtime"] != pdf_path.stat().st_mtime:
                return None
            # Migrar: reescribir con hash para que futuras cargas no dependan del mtime
            data["hash"] = pdf_hash
            data.pop("mtime", None)
            with open(cache_path, "w", encoding="utf-8") as fw:
                json.dump(data, fw, ensure_ascii=False)
        else:
            return None
        if data.get("zoom") != _OCR_ZOOM:
            return None
        # Invalidar si saldo_final cambió (afecta los movimientos de recuperación)
        cached_sf = data.get("saldo_final")
        if saldo_final is not None and cached_sf != saldo_final:
            return None
        movs: list[BancoRecord] = []
        for m in data["movs"]:
            rec = dict(m)
            rec["fecha"] = date.fromisoformat(m["fecha"]) if m["fecha"] else None
            movs.append(rec)
        return movs, data.get("saldo_ocr")
    except Exception:
        return None


def _save_ocr_cache(
    pdf_path: Path,
    movs: list[BancoRecord],
    saldo_ocr: Optional[float],
    saldo_final: Optional[float] = None,
) -> None:
    import json
    cache_path = _ocr_cache_path(pdf_path)
    try:
        serializable = []
        for m in movs:
            rec = dict(m)
            rec["fecha"] = m["fecha"].isoformat() if m["fecha"] else None
            serializable.append(rec)
        with open(cache_path, "w", encoding="utf-8") as f:
            json.dump({
                "hash": _pdf_hash(pdf_path),
                "zoom": _OCR_ZOOM,
                "saldo_ocr": saldo_ocr,
                "saldo_final": saldo_final,
                "movs": serializable,
            }, f, ensure_ascii=False)
    except Exception:
        pass


def parse_extracto_ocr(
    path: Path, stop_marker: str, saldo_final: Optional[float] = None
) -> tuple[list[BancoRecord], Optional[float]]:
    """Parser OCR para extractos bancarios en formato imagen (PDF escaneado).

    Renderiza cada página con PyMuPDF (3x), aplica EasyOCR y reconstruye
    movimientos usando las mismas reglas de columnas que el parser pdfplumber.
    También captura el saldo de cierre de la línea del marcador de corte,
    evitando una segunda pasada OCR.

    Usa un archivo de cache ``<pdf>.ocr_cache.json`` para evitar repetir el OCR
    si el PDF no cambió desde la última ejecución.

    Returns:
        Tupla ``(movimientos, saldo_cierre)``. ``saldo_cierre`` es ``None``
        si no se encontró la línea del marcador.
    """
    cached = _load_ocr_cache(path, saldo_final)
    if cached is not None:
        return cached

    try:
        import fitz        # PyMuPDF
        import numpy as np
    except ImportError:
        return [], None
    try:
        reader = _get_ocr_reader()
    except Exception:
        return [], None

    movs: list[BancoRecord] = []
    saldo_ocr: Optional[float] = None
    stop_clean = stop_marker.replace(" ", "").upper()

    with fitz.open(str(path)) as doc:
        cur_date: Optional[date] = None
        stop = False

        # Acumuladores globales para recuperación cross-page
        all_movs_with_y: list[tuple[int, dict]] = []
        all_saldo_checks: list[tuple[int, float]] = []
        _PAGE_OFFSET = 30_000  # separación de espacio y entre páginas (3x: ~2400px/página)

        for page_idx, page in enumerate(doc):
            if stop:
                break
            pix = page.get_pixmap(
                matrix=fitz.Matrix(_OCR_ZOOM, _OCR_ZOOM),
                colorspace=fitz.csGRAY,
            )
            img = np.frombuffer(pix.samples, dtype=np.uint8).reshape(
                pix.height, pix.width
            )
            results = reader.readtext(img, detail=1, paragraph=False)

            # Agrupar tokens por línea (buckets de 10 px para no fusionar filas contiguas)
            line_map: dict[int, list] = {}
            for (bbox, text, _conf) in results:
                x0 = min(p[0] for p in bbox)
                x1 = max(p[0] for p in bbox)
                y0 = min(p[1] for p in bbox)
                key = round(y0 / 10) * 10
                line_map.setdefault(key, []).append((x0, x1, y0, text.strip()))

            # Variante del marcador sin "/" para tolerar que OCR lo lea mal ("|", "1", etc.)
            stop_clean_alt = re.sub(r"[^A-Z0-9]", "", stop_marker.upper())

            for key in sorted(line_map):
                toks = sorted(line_map[key], key=lambda t: t[0])
                linea_clean = "".join(t[3] for t in toks).replace(" ", "").upper()
                linea_alphanum = re.sub(r"[^A-Z0-9]", "", linea_clean)
                linea_desc = " ".join(t[3] for t in toks if t[0] < _AMT_X0)

                if stop_clean in linea_clean or stop_clean_alt in linea_alphanum:
                    # Capturar saldo de cierre del marcador de corte
                    all_frags = [t[3] for t in toks]
                    amt = _normalize_ocr_amount(all_frags)
                    if amt is not None:
                        saldo_ocr = amt
                    else:
                        # Intentar token a token (el saldo puede estar separado)
                        for t in reversed(toks):
                            a = _normalize_ocr_amount([t[3]])
                            if a is not None:
                                saldo_ocr = a
                                break
                    stop = True
                    break

                # Saltar líneas de encabezado/pie de página (no son movimientos)
                if _SKIP_LINE_RE.search(linea_desc):
                    continue

                # Fecha
                date_toks = [t for t in toks if t[0] < _DATE_X0]
                d = _reconstruct_ocr_date(date_toks)
                if d:
                    cur_date = d

                # Saldo corriente (columna derecha): capturar para recuperación
                saldo_frags = [t[3] for t in toks
                               if t[0] > _SALDO_X0_WIDE
                               and not _DATE_TOKEN_RE.match(t[3].strip())
                               and any(c.isdigit() for c in t[3])]
                saldo_amt = _normalize_ocr_amount(saldo_frags)
                # Umbral de 5,000 para filtrar fragmentos de OCR misleídos (ej: 783.27)
                if saldo_amt and saldo_amt >= 5_000:
                    global_key = page_idx * _PAGE_OFFSET + key
                    all_saldo_checks.append((global_key, saldo_amt))

                # Importes: fragmentos en zona DEBITO / CREDITO
                # Excluir tokens con formato de fecha (ej: "30/09/2025" en col. saldo)
                deb_frags = [t[3] for t in toks
                             if _AMT_X0 < t[0] <= _DEB_X1
                             and not _DATE_TOKEN_RE.match(t[3].strip())]
                cred_frags = [t[3] for t in toks
                              if _DEB_X1 < t[0] <= _CRED_X1
                              and not _DATE_TOKEN_RE.match(t[3].strip())]

                deb = _normalize_ocr_amount(deb_frags)
                cred = _normalize_ocr_amount(cred_frags)

                if deb is None and cred is None:
                    continue
                if cur_date is None:
                    continue

                # Comprobante (token numérico al inicio de COMBTE)
                combte_toks = [t[3] for t in toks if _COMBTE_X0 <= t[0] < _COMBTE_X1]
                combte = ""
                if combte_toks and re.fullmatch(r"\d{3,}", combte_toks[0]):
                    combte = combte_toks[0]

                # Descripción
                desc_toks = [t[3] for t in toks if _COMBTE_X1 <= t[0] < _AMT_X0]
                desc = " ".join(desc_toks)

                gkey = page_idx * _PAGE_OFFSET + key
                if deb is not None:
                    rec: BancoRecord = {
                        "fecha": cur_date, "combte": combte, "desc": desc,
                        "tipo": "D", "importe": deb,
                    }
                    movs.append(rec)
                    all_movs_with_y.append((gkey, rec))
                if cred is not None:
                    rec = {
                        "fecha": cur_date, "combte": combte, "desc": desc,
                        "tipo": "C", "importe": cred,
                    }
                    movs.append(rec)
                    all_movs_with_y.append((gkey, rec))

        # Anclar con saldo de cierre: evita que los últimos segmentos generen DEBs
        # fantasma que el statement recupera al 98.6M real al final.
        anchor = saldo_final if saldo_final is not None else saldo_ocr
        if anchor is not None and all_saldo_checks:
            last_key = max(k for k, _ in all_saldo_checks)
            all_saldo_checks.append((last_key + 1, anchor))

        # Recuperación cross-page: un único pase con todos los checkpoints globales
        filtered_checks = _filter_partial_saldo(all_saldo_checks)
        recovered = _recover_from_saldo(all_movs_with_y, filtered_checks, cur_date)
        movs.extend(recovered)

    movs = _filter_sircreb_outliers(movs)
    _save_ocr_cache(path, movs, saldo_ocr, saldo_final)
    return movs, saldo_ocr


def is_scanned_pdf(path: Path) -> bool:
    """Devuelve True si el PDF no tiene texto seleccionable (es imagen/escaneado)."""
    with pdfplumber.open(str(path)) as pdf:
        for page in pdf.pages[:3]:
            if page.extract_words(use_text_flow=False):
                return False
    return True


def parse_extracto(
    path: Path, stop_marker: str, saldo_final: Optional[float] = None
) -> tuple[list[BancoRecord], bool]:
    """Extrae los movimientos del extracto bancario en PDF.

    Intenta primero con pdfplumber (PDFs digitales con texto seleccionable).
    Si no extrae ningún movimiento, cae automáticamente al parser OCR
    (``parse_extracto_ocr``) para PDFs escaneados/imagen.

    Returns:
        Tupla ``(movimientos, ocr_usado)``.
    """
    # Verificar cache OCR antes de abrir pdfplumber: evita iterar 20+ páginas
    # bezier (costoso) cuando el resultado ya está guardado.
    cached = _load_ocr_cache(path, saldo_final)
    if cached is not None:
        movs, saldo_ocr = cached
        movs = _filter_sircreb_outliers(movs)
        return movs, True, saldo_ocr

    movs: list[BancoRecord] = []
    with pdfplumber.open(str(path)) as pdf:
        cur_date: Optional[date] = None
        stop = False
        for page in pdf.pages:
            if stop:
                break
            words = page.extract_words(use_text_flow=False)
            lines: dict[int, list] = {}
            for w in words:
                key = round(w["top"] / 3.0)
                lines.setdefault(key, []).append(w)
            for key in sorted(lines):
                toks = sorted(lines[key], key=lambda w: w["x0"])
                texts = [t["text"] for t in toks]
                linea = " ".join(texts)
                if stop_marker in linea:
                    stop = True
                    break
                d = None
                if texts and DATE_RE.match(texts[0]):
                    dd, mm, yy = texts[0].split("/")
                    d = date(2000 + int(yy), int(mm), int(dd))
                    cur_date = d
                deb = cred = None
                for t in toks:
                    if AMOUNT_RE.match(t["text"]):
                        x1 = t["x1"]
                        if x1 < 450:
                            deb = parse_importe(t["text"])
                        elif x1 < 555:
                            cred = parse_importe(t["text"])
                if d is None and deb is None and cred is None:
                    continue
                if deb is None and cred is None:
                    continue
                desc_toks = [
                    t["text"] for t in toks
                    if not AMOUNT_RE.match(t["text"]) and not DATE_RE.match(t["text"])
                ]
                combte = ""
                if desc_toks and re.fullmatch(r"\d{3,}", desc_toks[0]):
                    combte = desc_toks[0]
                    desc_toks = desc_toks[1:]
                desc = " ".join(desc_toks).strip()
                if deb is not None:
                    movs.append({"fecha": cur_date, "combte": combte,
                                 "desc": desc, "tipo": "D", "importe": deb})
                if cred is not None:
                    movs.append({"fecha": cur_date, "combte": combte,
                                 "desc": desc, "tipo": "C", "importe": cred})

    # Fallback OCR para PDFs escaneados (sin texto seleccionable).
    # parse_extracto_ocr captura también el saldo de cierre en el mismo pase.
    if not movs:
        movs, saldo_ocr = parse_extracto_ocr(path, stop_marker, saldo_final=saldo_final)
        return movs, True, saldo_ocr

    return movs, False, None


def saldo_final_pdf(path: Path, stop_marker: str) -> Optional[float]:
    """Devuelve el saldo bancario de cierre leyendo el PDF con pdfplumber."""
    saldo: Optional[float] = None
    with pdfplumber.open(str(path)) as pdf:
        for page in pdf.pages:
            words = page.extract_words(use_text_flow=False)
            lines: dict[int, list] = {}
            for w in words:
                lines.setdefault(round(w["top"] / 3.0), []).append(w)
            for key in sorted(lines):
                toks = sorted(lines[key], key=lambda w: w["x0"])
                linea = " ".join(t["text"] for t in toks)
                amts = [parse_importe(t["text"]) for t in toks if AMOUNT_RE.match(t["text"])]
                if stop_marker in linea and amts:
                    saldo = amts[-1]
    return saldo


# ===========================================================================
# 2. Parseo de la CONTABILIDAD (Excel exportado de SAP)
# ===========================================================================
def parse_contabilidad(path: Path) -> list[ContaRecord]:
    """Extrae los movimientos del libro contable exportado de SAP.

    Lee columnas por posición: fecha, clase, número, debe, haber, cuit,
    sujeto, descripción y subdiario. El lado bancario esperado se deriva del
    debe/haber (debe > 0 → crédito bancario).
    """
    wb = openpyxl.load_workbook(str(path), data_only=True)
    ws = wb.active
    rows: list[ContaRecord] = []
    for i, r in enumerate(ws.iter_rows(values_only=True)):
        if i == 0:
            continue
        if r[0] is None:
            continue
        fec = r[0].date() if isinstance(r[0], datetime) else r[0]
        debe = float(r[5]) if r[5] else 0.0
        haber = float(r[6]) if r[6] else 0.0
        rows.append({
            "fecha": fec, "clase": r[3], "numero": r[4],
            "debe": debe, "haber": haber,
            "cuit": r[7], "sujeto": r[8], "desc": r[9], "subdiario": r[10],
            "tipo_bco": "C" if debe > 0 else "D",
            "importe": debe if debe > 0 else haber,
        })
    return rows


# ===========================================================================
# 3. Anular autocancelatorios (mismo importe, lados opuestos, misma fecha)
# ===========================================================================
def marcar_anulados(
    records: list, side_of: Callable[[dict], str]
) -> list[tuple[int, int]]:
    """Marca como anulados los pares que salen y vuelven (netean entre sí).

    Args:
        records: Lista de movimientos (contables o bancarios).
        side_of: Función que devuelve ``"IN"`` o ``"OUT"`` para cada registro.

    Returns:
        Lista de pares de índices ``(salida, entrada)`` anulados.
    """
    by: dict[tuple, dict[str, list[int]]] = defaultdict(
        lambda: {"IN": [], "OUT": []}
    )
    for k, r in enumerate(records):
        r["anulado"] = False
        by[(round(r["importe"], 2), r["fecha"])][side_of(r)].append(k)
    pares: list[tuple[int, int]] = []
    for d in by.values():
        for a, b in zip(d["OUT"], d["IN"]):
            records[a]["anulado"] = records[b]["anulado"] = True
            pares.append((a, b))
    return pares


# Tolerancia asimétrica de fecha: créditos ±5 días, débitos ±10 días.
# Los débitos tienen ventana mayor porque los pagos de importación suelen
# registrarse en SAP varios días después de que el banco los debita.
_DATE_TOL_C = 5
_DATE_TOL_D = 10


# ===========================================================================
# 4. Matching uno-a-uno (importe + lado invertido, fecha como desempate)
# ===========================================================================
def conciliar(
    conta: list[ContaRecord], banco: list[BancoRecord]
) -> tuple[list[bool], list[bool], list[tuple[int, int]]]:
    """Concilia contabilidad y banco uno-a-uno.

    Empareja por (lado bancario, importe) filtrando candidatos dentro de la
    ventana de fechas y eligiendo el de fecha más cercana entre los válidos.
    Usa tolerancia asimétrica: 5 días para créditos, 10 días para débitos.

    Returns:
        ``(conta_match, banco_match, pares)`` con los flags de conciliación y
        los pares de índices emparejados.
    """
    idx: dict[tuple, list[int]] = {}
    for j, b in enumerate(banco):
        if b.get("anulado"):
            continue
        idx.setdefault((b["tipo"], round(b["importe"], 2)), []).append(j)
    banco_match = [False] * len(banco)
    conta_match = [False] * len(conta)
    pares: list[tuple[int, int]] = []
    for i in sorted(range(len(conta)), key=lambda k: conta[k]["fecha"]):
        c = conta[i]
        if c.get("anulado"):
            continue
        cands = idx.get((c["tipo_bco"], round(c["importe"], 2)), [])
        libres = [j for j in cands if not banco_match[j]]
        if not libres:
            continue
        dtol = _DATE_TOL_C if c["tipo_bco"] == "C" else _DATE_TOL_D
        libres = [
            j for j in libres
            if banco[j]["fecha"] and c["fecha"]
            and abs((banco[j]["fecha"] - c["fecha"]).days) <= dtol
        ]
        if not libres:
            continue
        libres.sort(key=lambda j: abs((banco[j]["fecha"] - c["fecha"]).days))
        j = libres[0]
        banco_match[j] = True
        conta_match[i] = True
        pares.append((i, j))
    return conta_match, banco_match, pares


# ===========================================================================
# 5. Matching AGRUPADO (N:1) por suma
# ===========================================================================
def _best_group(
    pool: list[tuple[int, int]], target: int, banco: list[BancoRecord],
    max_k: int = 8, max_span: int = 3,
) -> Optional[list[int]]:
    """Busca el mejor subconjunto del pool cuya suma iguale ``target``.

    Prioriza el menor span de fechas y la menor cantidad de elementos.
    Los importes se manejan en centavos (enteros) para evitar errores de
    punto flotante.
    """
    n = len(pool)
    if n < 2:
        return None
    best = None
    for k in range(2, min(max_k, n) + 1):
        for comb in combinations(range(n), k):
            if sum(pool[t][1] for t in comb) != target:
                continue
            fechas = [banco[pool[t][0]]["fecha"] for t in comb]
            span = (max(fechas) - min(fechas)).days
            if span > max_span:
                continue
            cand = (span, k, [pool[t][0] for t in comb])
            if best is None or (span, k) < (best[0], best[1]):
                best = cand
        if best is not None and best[0] == 0:
            break
    return best[2] if best else None


def conciliar_grupos(
    conta: list[ContaRecord], banco: list[BancoRecord],
    conta_match: list, banco_match: list[bool],
    win: int = 4, cap: int = 22,
) -> dict[int, list[int]]:
    """Concilia N movimientos de banco contra 1 de contabilidad por suma.

    Para cada movimiento contable pendiente, arma un pool de movimientos
    bancarios libres del mismo lado, categoría "Operacion" y dentro de una
    ventana de fechas, y busca la combinación que sume el importe objetivo.
    Muta ``conta_match`` (marca ``"GRUPO"``) y ``banco_match``.

    Returns:
        Mapa ``{índice_conta: [índices_banco]}`` de los grupos encontrados.
    """
    cents = lambda x: int(round(x * 100))
    grupos: dict[int, list[int]] = {}
    for i, c in enumerate(conta):
        if conta_match[i] or c.get("anulado"):
            continue
        side, tgt = c["tipo_bco"], cents(c["importe"])
        pool = [(j, cents(banco[j]["importe"])) for j in range(len(banco))
                if not banco_match[j] and not banco[j].get("anulado")
                and banco[j]["tipo"] == side
                and categoria_banco(banco[j]["desc"]) == CATEGORIA_OPERACION
                and banco[j]["fecha"]
                and abs((banco[j]["fecha"] - c["fecha"]).days) <= win]
        pool.sort(key=lambda p: abs((banco[p[0]]["fecha"] - c["fecha"]).days))
        pool = pool[:cap]
        sel = _best_group(pool, tgt, banco)
        if sel:
            grupos[i] = sel
            for j in sel:
                banco_match[j] = True
            conta_match[i] = "GRUPO"
    return grupos
