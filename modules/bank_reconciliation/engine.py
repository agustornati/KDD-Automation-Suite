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
from typing import Callable, Optional

import openpyxl
import pdfplumber

from .models import BancoRecord, ContaRecord
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
def parse_extracto(path: Path, stop_marker: str) -> list[BancoRecord]:
    """Extrae los movimientos del extracto bancario en PDF.

    Recorre el PDF línea por línea agrupando palabras por su coordenada
    vertical y clasifica los importes en débito/crédito según su posición
    horizontal. Se detiene al encontrar la línea de corte (``stop_marker``).
    """
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
    return movs


def saldo_final_pdf(path: Path, stop_marker: str) -> Optional[float]:
    """Devuelve el saldo bancario de cierre.

    Es el último importe de la línea ``SALDO AL dd/mm`` del extracto.
    """
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


# ===========================================================================
# 4. Matching uno-a-uno (importe + lado invertido, fecha como desempate)
# ===========================================================================
def conciliar(
    conta: list[ContaRecord], banco: list[BancoRecord]
) -> tuple[list[bool], list[bool], list[tuple[int, int]]]:
    """Concilia contabilidad y banco uno-a-uno.

    Empareja por (lado bancario, importe) eligiendo, entre los candidatos
    libres, el de fecha más cercana.

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
        libres.sort(key=lambda j: abs((banco[j]["fecha"] - c["fecha"]).days)
                    if banco[j]["fecha"] and c["fecha"] else 999)
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
