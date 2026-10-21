"""Aannemer section: pure parsing/matching logic (no Streamlit).

Compares a contractor's PDF offerte (e.g. Van Wijnen 'Afbouw', structurally a
printed Excel bidsheet — same column layout as Koeling's 'Offertetemplate
KT') against the IB workbook's Budget sheet. Unlike Koeling, the Budget rows
for this section have no numeric-prefix bridge to join against Art.Nr.Jumbo
(see match_aannemer_budget), and the offerte's own Art.Nr.Jumbo column is
already filled in by the leverancier rather than needing to be backfilled.

Split out of pages/3_🧱_Aannemer.py so it can be unit-tested without a
Streamlit script-run context — importing that page module executes
st.set_page_config()/session_state code at import time.
"""

import io
import re
from dataclasses import dataclass
from pathlib import Path
from typing import Dict, List, Optional

import pandas as pd
import pdfplumber

from matching_koeling import (
    HAS_RAPIDFUZZ,
    _budget_code_totals,
    _norm_text,
    budget_ib_only_rows,
    budget_leaf_rows,
    build_categorie_totalen,
    load_budget_df,
)
from vergelijking import (
    OfferteVergelijking,
    VergelijkingRegel,
    categorieen_in_volgorde,
    leverancier_uit_rijen,
    match_opmerking,
    met_ib_only_regels,
)

if HAS_RAPIDFUZZ:
    from rapidfuzz import fuzz

__all__ = [
    "HAS_RAPIDFUZZ",
    "BUDGET_AANNEMER_VANWIJNEN_ROW_END",
    "BUDGET_AANNEMER_VANWIJNEN_ROW_START",
    "BudgetMatch",
    "OfferteItem",
    "build_aannemer_df",
    "build_aannemer_vergelijking",
    "build_categorie_totalen",
    "get_aannemer_totaal_ib",
    "load_art_nr_mapping",
    "load_budget_df",
    "load_categorie_budget_codes",
    "match_aannemer_budget",
    "parse_excel",
    "parse_pdf",
    "read_offerte",
]


@dataclass
class OfferteItem:
    row: int
    pos: Optional[str]  # "Pos 1 Kozijnen en deuren", ... or "NRFP"
    art_nr_jumbo: Optional[str]
    manual_nr: Optional[str]
    art_nr_leverancier: Optional[str]
    omschrijving: str
    toelichting: Optional[str]
    eenheid: Optional[str]
    aantal: Optional[float]
    materiaalkosten_pe: Optional[float]
    normtijd_min_pe: Optional[float]
    tarief_uur: Optional[float]
    prijs_pe: Optional[float]
    totale_uren: Optional[float]
    totale_materiaalkosten: Optional[float]
    arbeid: Optional[float]
    totaal: Optional[float]
    opmerkingen: Optional[str]
    page: int = 0


@dataclass
class BudgetMatch:
    budget_row: int
    code: Optional[str]
    omschrijving: str
    aantal: Optional[float]
    prijs: Optional[float]
    totaal: Optional[float]
    score: float
    source: str  # "handmatig" | "fuzzy"


#: Budget-sheet rows covering the Aannemer (Van Wijnen) 'A.02' section.
#: Baked-in leverancier name since "Aannemer" is a role, not one supplier —
#: the IB workbook also has an 'Aannemer van de Laar' sheet for another one.
#: Row 273 (the range's own start row) is the 'A.02 Aannemerswerk' section
#: header itself — its own Totaal is the authoritative whole-section IB
#: figure (see get_aannemer_totaal_ib), since summing the per-Pos categorie
#: rollup misses budget-only subheadings that have no offerte Pos at all
#: (Reiskosten, Overig).
BUDGET_AANNEMER_VANWIJNEN_ROW_START = 273
BUDGET_AANNEMER_VANWIJNEN_ROW_END = 456


def get_aannemer_totaal_ib(
    bdf: Optional[pd.DataFrame],
    row_start: int = BUDGET_AANNEMER_VANWIJNEN_ROW_START, row_end: int = BUDGET_AANNEMER_VANWIJNEN_ROW_END,
) -> Optional[float]:
    """Sum every A.02.0X subheading's own Totaal directly (Kozijnen, Wanden,
    ... Reiskosten, Overig — all 12, not just the ones with a matching
    offerte Pos) — the whole-section IB figure for the Samenvatting's top
    'Totaal IB' KPI, independent of build_categorie_totalen's per-Pos
    rollup below (which only covers categories the offerte itself quotes).
    Equals the 'A.02 Aannemerswerk' row's own Totaal, computed explicitly
    from its children rather than trusting that single cell."""
    if bdf is None:
        return None
    code_totals = _budget_code_totals(bdf, row_start, row_end)
    total = 0.0
    found = False
    for code, (_desc, _aantal, totaal) in code_totals.items():
        if not isinstance(code, str) or not code.startswith("A.02.") or code.count(".") != 2:
            continue
        if totaal:
            total += totaal
            found = True
    return total if found else None


_WS_RE = re.compile(r"\s+")


def _dutch_number(s) -> Optional[float]:
    """'€ 1 07.484,15' -> 107484.15 — pdfplumber's word-clustering inserts a
    stray space *inside* digit groups for this PDF's font (verified against
    the real file), so all internal whitespace must be stripped, not just
    trimmed, before applying Dutch decimal conversion ('.' thousands, ','
    decimal). Blank cells render as a bare '-'. A real number (an Excel
    cell) passes through as-is — '84.49' as text would otherwise read 8449."""
    if s is None:
        return None
    if isinstance(s, (int, float)):
        return None if pd.isna(s) else float(s)
    s = _WS_RE.sub("", str(s)).replace("€", "")
    if s in ("", "-"):
        return None
    try:
        return float(s.replace(".", "").replace(",", "."))
    except ValueError:
        return None


_HEADER_COL = {
    "art_nr_jumbo": 0, "manual_nr": 1, "art_nr_leverancier": 2, "pos": 3,
    "omschrijving": 4, "toelichting": 5, "eenheid": 6, "aantal": 7,
    "materiaalkosten_pe": 8, "normtijd_min_pe": 9, "tarief_uur": 10,
    "prijs_pe": 11, "totale_uren": 12, "totale_materiaalkosten": 13,
    "arbeid": 14, "totaal": 15, "opmerkingen": 16,
}


def _is_header_row(raw_row: list) -> bool:
    # The Excel header reads 'Art. nr.  Jumbo' (double space), the PDF's
    # 'Art. nr. Jumbo' — compare with whitespace collapsed.
    return bool(raw_row) and _WS_RE.sub(" ", str(raw_row[0] or "")).strip() == "Art. nr. Jumbo"


def _parse_table_row(raw_row: list, row_num: int, page: int) -> Optional[OfferteItem]:
    """Turn one pdfplumber extract_tables() row into an OfferteItem, or None
    for a header/blank row. Column indices are fixed (_HEADER_COL) — the
    17-column layout is constant across every Pos-block and the NRFP block
    alike (verified against the real PDF), so no per-section header_map
    rebuild is needed."""
    def cell(key):
        idx = _HEADER_COL[key]
        return raw_row[idx] if idx < len(raw_row) else None

    omschrijving = cell("omschrijving")
    if not omschrijving or not str(omschrijving).strip():
        return None

    art_nr = cell("art_nr_jumbo")
    return OfferteItem(
        row=row_num,
        pos=str(cell("pos")).strip() if cell("pos") else None,
        art_nr_jumbo=str(art_nr).strip() if art_nr else None,
        manual_nr=str(cell("manual_nr")).strip() if cell("manual_nr") else None,
        art_nr_leverancier=str(cell("art_nr_leverancier")).strip() if cell("art_nr_leverancier") else None,
        omschrijving=str(omschrijving).strip(),
        toelichting=str(cell("toelichting")).strip() if cell("toelichting") else None,
        eenheid=str(cell("eenheid")).strip() if cell("eenheid") else None,
        aantal=_dutch_number(cell("aantal")),
        materiaalkosten_pe=_dutch_number(cell("materiaalkosten_pe")),
        normtijd_min_pe=_dutch_number(cell("normtijd_min_pe")),
        tarief_uur=_dutch_number(cell("tarief_uur")),
        prijs_pe=_dutch_number(cell("prijs_pe")),
        totale_uren=_dutch_number(cell("totale_uren")),
        totale_materiaalkosten=_dutch_number(cell("totale_materiaalkosten")),
        arbeid=_dutch_number(cell("arbeid")),
        totaal=_dutch_number(cell("totaal")),
        opmerkingen=str(cell("opmerkingen")).strip() if cell("opmerkingen") else None,
        page=page,
    )


def parse_pdf(file_bytes: bytes) -> List[OfferteItem]:
    """Each page also contains small metadata/summary tables (project info
    box, totals box) alongside the real line-item table — only tables that
    contain at least one 'Art. nr. Jumbo' header row are the 17-column
    line-item tables this parser targets; anything else is skipped
    entirely, since its column layout doesn't match _HEADER_COL at all."""
    items: List[OfferteItem] = []
    row_num = 0
    with pdfplumber.open(io.BytesIO(file_bytes)) as pdf:
        for page_num, page in enumerate(pdf.pages, start=1):
            for table in page.extract_tables():
                if not any(_is_header_row(r) for r in table):
                    continue
                for raw_row in table:
                    if _is_header_row(raw_row):
                        continue
                    row_num += 1
                    item = _parse_table_row(raw_row, row_num, page_num)
                    if item is not None:
                        items.append(item)
    return items


#: 0-based column of 'Totale waarde offerte' in the offerte layout — used to
#: tell the 'vertaling bidsheet - offerte' sheet apart from the workbook's
#: 'Bidsheet' (RFP) sheet, which has its own 'Art. nr. Jumbo' header rows but
#: a different column layout.
_COL_TOTAAL_HEADER = _HEADER_COL["totaal"]


def _excel_rows(file_bytes: bytes):
    """(row number, cells) of the workbook's offerte sheet — the sheet whose
    'Art. nr. Jumbo' header row has 'Totale waarde offerte' in the Totaal
    column, preferring a sheet named '…offerte' — with NaN cells as None.
    Empty when no sheet has that layout."""
    sheets = pd.read_excel(io.BytesIO(file_bytes), sheet_name=None, header=None, engine="openpyxl")
    ordered = sorted(sheets.items(), key=lambda kv: "offerte" not in kv[0].lower())
    for _name, df in ordered:
        rows = [[None if (isinstance(v, float) and pd.isna(v)) else v for v in raw] for raw in df.values.tolist()]
        for raw in rows:
            if _is_header_row(raw) and len(raw) > _COL_TOTAAL_HEADER and                     _WS_RE.sub(" ", str(raw[_COL_TOTAAL_HEADER] or "")).strip().lower() == "totale waarde offerte":
                return list(enumerate(rows, start=1))
    return []


def parse_excel(file_bytes: bytes) -> List[OfferteItem]:
    """Same as parse_pdf, for the Excel the PDF is printed from (sheet
    'vertaling bidsheet - offerte'): its 17 columns sit at the same
    positions as _HEADER_COL (verified against Van de Laar's Afbouw
    workbooks), so every row after the first header row goes through the
    same _parse_table_row. The summary block above the first header row is
    skipped."""
    items: List[OfferteItem] = []
    seen_header = False
    for r, raw in _excel_rows(file_bytes):
        if _is_header_row(raw):
            seen_header = True
            continue
        if not seen_header:
            continue
        item = _parse_table_row(raw, r, page=0)
        if item is not None:
            items.append(item)
    return items


def _pdf_table_rows(file_bytes: bytes):
    r = 0
    with pdfplumber.open(io.BytesIO(file_bytes)) as pdf:
        for page in pdf.pages:
            for table in page.extract_tables():
                for raw in table:
                    r += 1
                    yield r, raw


def read_offerte(file_bytes: bytes, filename: str) -> tuple:
    """(offerte items, leverancier) from the offerte as PDF or as Excel —
    the leverancier from its 'Prijsopgave Leverancier' line, else 'Offerte'."""
    if filename.lower().endswith(".pdf"):
        items, rows = parse_pdf(file_bytes), _pdf_table_rows(file_bytes)
    else:
        items, rows = parse_excel(file_bytes), _excel_rows(file_bytes)
    return items, leverancier_uit_rijen(rows) or "Offerte"


#: Persistent Art.Nr.Jumbo -> canonical Budget omschrijving lookup, keyed the
#: opposite way from Koeling's mapping: this offerte already has real
#: Art.Nr.Jumbo values, so once a fuzzy match is confirmed correct by a
#: human, it's recorded here as an exact override for future runs.
ART_NR_MAPPING_DIR = Path(__file__).parent / "data"


def load_art_nr_mapping(leverancier: str = "vanwijnen") -> Dict[str, str]:
    path = ART_NR_MAPPING_DIR / f"aannemer_art_nr_mapping_{leverancier}.csv"
    if not path.exists():
        return {}
    df = pd.read_csv(path)
    return {str(row["art_nr"]).strip(): str(row["budget_omschrijving"]).strip() for _, row in df.iterrows()}


def load_categorie_budget_codes() -> Dict[str, List[tuple]]:
    """Read data/Aanemer_categorie_budget_codes.csv: offerte 'Pos N ...' label
    -> [(budget_code, budget_omschrijving), ...]. Keyed by the offerte's own
    raw Pos string (not a normalized/stripped label) so it matches
    OfferteItem.pos verbatim — same convention as Koeling's
    load_categorie_budget_codes, whose result feeds directly into
    build_categorie_totalen (reused as-is, see matching_koeling.py — it's
    already generic over the mapping/row range, no Koeling-specific logic)."""
    path = ART_NR_MAPPING_DIR / "Aanemer_categorie_budget_codes.csv"
    if not path.exists():
        return {}
    df = pd.read_csv(path)
    mapping: Dict[str, List[tuple]] = {}
    for _, row in df.iterrows():
        categorie = str(row["categorie"]).strip()
        code = str(row["budget_code"]).strip()
        omschrijving = row.get("budget_omschrijving")
        omschrijving = str(omschrijving).strip() if pd.notna(omschrijving) else None
        mapping.setdefault(categorie, []).append((code, omschrijving))
    return mapping


def match_aannemer_budget(
    offerte: List[OfferteItem], bdf: Optional[pd.DataFrame], mapping: Dict[str, str],
    row_start: int = BUDGET_AANNEMER_VANWIJNEN_ROW_START, row_end: int = BUDGET_AANNEMER_VANWIJNEN_ROW_END,
    fuzzy_threshold: int = 40,
) -> Dict[int, BudgetMatch]:
    """Offerte-row -> Budget-row join. Unlike Koeling's build_budget_link,
    Aannemer's Budget rows carry no numeric prefix in their omschrijving to
    bridge against Art.Nr.Jumbo, so the join is text-based: an exact CSV
    override first (mapping[art_nr] == Budget omschrijving), then rapidfuzz
    description matching for the rest."""
    if bdf is None:
        return {}

    # Only leaf item rows ('A.02.04.01') — matching against a subtotal row
    # ('A.02.04', 'A.02') makes NRFP/offerte rows fuzzy-match things like
    # "Algemeen"/"Reiskosten".
    candidates = budget_leaf_rows(bdf, row_start, row_end)

    used_budget = set()
    matches: Dict[int, BudgetMatch] = {}

    def _claim(i, c, score, source):
        used_budget.add(c["row"])
        matches[i] = BudgetMatch(
            budget_row=c["row"], code=c["code"], omschrijving=c["desc"], aantal=c["aantal"],
            prijs=c["prijs"], totaal=c["totaal"], score=score, source=source,
        )

    # NRFP rows are extras outside the RFP scope (huur rijplaten,
    # precariokosten, ...) — they don't correspond to any A.02 Budget leaf
    # item, so attempting to match them just produces nonsense fuzzy hits
    # (verified against the real file: NRFP rows scored against unrelated
    # Budget lines like "Reiskosten"/"Algemeen"). Same treatment as
    # Koeling's NRFP sheet: kept entirely out of the Budget join, shown
    # separately in the Samenvatting instead.
    matchable = [(i, o) for i, o in enumerate(offerte) if o.pos != "NRFP"]

    # Pass 1: exact CSV override — takes priority, mirrors Koeling's
    # manual-before-fuzzy call order.
    for i, o in matchable:
        if not o.art_nr_jumbo:
            continue
        target_desc = mapping.get(o.art_nr_jumbo)
        if not target_desc:
            continue
        for c in candidates:
            if c["row"] not in used_budget and c["desc"] == target_desc:
                _claim(i, c, 100.0, "handmatig")
                break

    # Pass 2: rapidfuzz description fallback for whatever's left.
    if HAS_RAPIDFUZZ:
        pairs = []
        for i, o in matchable:
            if i in matches or not o.omschrijving:
                continue
            for c in candidates:
                if c["row"] in used_budget:
                    continue
                score = fuzz.token_set_ratio(_norm_text(o.omschrijving), _norm_text(c["desc"]))
                if score >= fuzzy_threshold:
                    pairs.append((score, i, c))
        pairs.sort(key=lambda x: -x[0])
        for score, i, c in pairs:
            if i in matches or c["row"] in used_budget:
                continue
            _claim(i, c, float(score), "fuzzy")

        # Fail-safe: a matched Budget row can still have no aantal — IB ended
        # up selecting a different, similarly-described variant than the one
        # matched here (e.g. a different gevelwand-type), which the CSV
        # override/exact-desc pass can't detect on its own. Re-point that
        # match to the closest remaining text match that does carry a real
        # aantal, so "Alle resultaten" doesn't show a blank IB Aantal when a
        # better candidate exists.
        for i, m in list(matches.items()):
            if m.aantal:
                continue
            o = offerte[i]
            if not o.omschrijving:
                continue
            best, best_score = None, fuzzy_threshold
            for c in candidates:
                if c["row"] in used_budget or not c["aantal"]:
                    continue
                score = fuzz.token_set_ratio(_norm_text(o.omschrijving), _norm_text(c["desc"]))
                if score > best_score:
                    best_score, best = score, c
            if best is not None:
                used_budget.discard(m.budget_row)
                _claim(i, best, float(best_score), "fuzzy")

    return matches


def build_aannemer_df(offerte: List[OfferteItem], matches: Dict[int, BudgetMatch]) -> pd.DataFrame:
    rows = []
    for i, o in enumerate(offerte):
        m = matches.get(i)

        aantal_o = o.aantal or 0
        aantal_i = m.aantal if (m and m.aantal is not None) else None
        aantal_diff = (aantal_o - aantal_i) if aantal_i is not None else None

        # o.prijs_pe is the PDF's own combined 'Prijs per eenheid (arbeid &
        # materiaal)' column — o.arbeid, despite its column header, is a
        # row-level *total* (arbeid_per_eenheid * aantal), not a per-unit
        # figure, so it must not be added again here (verified against the
        # real file: materiaalkosten_totaal + arbeid == totaal exactly).
        prijs_o = o.prijs_pe
        prijs_i = m.prijs if m else None
        prijs_diff = (prijs_o - prijs_i) if (prijs_o is not None and prijs_i is not None) else None
        prijs_diff_pct = (prijs_diff / prijs_i * 100) if (prijs_diff is not None and prijs_i) else None

        if aantal_diff is None or aantal_diff == 0:
            significant_aantal_diff = False
        elif aantal_i:
            significant_aantal_diff = abs(aantal_diff) / aantal_i * 100 > 10
        else:
            significant_aantal_diff = True

        if o.pos == "NRFP":
            status = "nrfp"
        elif m is None:
            status = "unmatched"
        elif prijs_diff is not None and abs(prijs_diff) > 0.01:
            status = "prijs_afwijking"
        elif significant_aantal_diff:
            status = "aantal_afwijking"
        else:
            status = "ok"

        # Recompute rather than trust the PDF's own 'Totale waarde offerte'
        # column verbatim, for consistency with Koeling's established
        # convention — even though this session's spot-check against the
        # real file found it matches exactly (prijs_pe * aantal == totaal).
        offerte_totaal = (prijs_o or 0) * aantal_o if aantal_o else None
        if m and m.totaal is not None:
            ib_totaal = m.totaal
        else:
            ib_totaal = (prijs_i or 0) * aantal_o if (aantal_o and prijs_i is not None) else None

        rows.append({
            "_status": status,
            "_offerte_totaal": offerte_totaal,
            "_ib_totaal": ib_totaal,
            "Pos": o.pos,
            "Offerte omschrijving": o.omschrijving,
            "IB omschrijving": m.omschrijving if m else None,
            "Match": m.source if m else "geen match",
            "Match %": f"{m.score:.0f}%" if m else "",
            "Eenheid": o.eenheid,
            "Offerte Aantal": aantal_o if aantal_o else None,
            "IB Aantal": aantal_i,
            "Aantal verschil": aantal_diff,
            "Offerte Prijs p.e.": prijs_o,
            "IB Prijs p.e.": prijs_i,
            "Prijs verschil (€)": prijs_diff,
            "Prijs verschil (%)": prijs_diff_pct,
        })
    return pd.DataFrame(rows)


def build_aannemer_vergelijking(
    offerte: List[OfferteItem], matches: Dict[int, BudgetMatch],
    categorie_totalen: Optional[Dict[str, float]] = None, naam: str = "Van Wijnen", bestand: str = "",
    bdf: Optional[pd.DataFrame] = None, categorie_mapping: Optional[Dict[str, List[tuple]]] = None,
    row_start: int = BUDGET_AANNEMER_VANWIJNEN_ROW_START, row_end: int = BUDGET_AANNEMER_VANWIJNEN_ROW_END,
) -> OfferteVergelijking:
    """The offerte next to the IB per line, for the shared 'Alle resultaten'
    layout (common.show_alle_resultaten) — lines with an aantal, NRFP left
    out (the Samenvatting shows it separately). Offerte prijs p.e. is the PDF/Excel's combined 'Prijs per eenheid'. IB totaal is the
    matched Budget row's own Totaal (0 when that row has none).

    With `bdf`, the Budget rows no shown line covers are added as IB-only
    lines (matching_koeling.budget_ib_only_rows, filed under their offerte
    categorie via `categorie_mapping`), so the IB Totaal adds up to the
    Samenvatting's Totaal IB."""
    regels, used = [], set()
    for i, o in enumerate(offerte):
        if o.pos == "NRFP" or not o.aantal:
            continue
        m = matches.get(i)
        if m:
            used.add(m.budget_row)
        regels.append(VergelijkingRegel(
            categorie=o.pos, art_nr_jumbo=o.art_nr_jumbo, omschrijving=o.omschrijving,
            ib_omschrijving=m.omschrijving if m else None, eenheid=o.eenheid,
            ib_aantal=m.aantal if m else None, ib_prijs=m.prijs if m else None,
            ib_totaal_budget=(m.totaal or 0.0) if m else None,
            offertes=[(o.aantal, o.prijs_pe)], opmerking=match_opmerking(m),
        ))
    if bdf is not None:
        code_naar_categorie = {code: cat for cat, entries in (categorie_mapping or {}).items() for code, _ in entries}
        ib_only = budget_ib_only_rows(bdf, row_start, row_end, "A.02.", used)
        regels = met_ib_only_regels(regels, ib_only, code_naar_categorie)
    return OfferteVergelijking(
        namen=[naam], bestanden=[bestand], regels=regels, nrfp=[[]],
        categorieen=categorieen_in_volgorde(regels), ib_categorie_totalen=dict(categorie_totalen or {}),
    )
