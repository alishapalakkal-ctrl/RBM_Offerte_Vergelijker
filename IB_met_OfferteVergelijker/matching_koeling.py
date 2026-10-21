"""Koeling section: pure parsing/matching logic (no Streamlit).

Split out of pages/2_🧊_Koeling.py so it can be unit-tested without a
Streamlit script-run context — importing that page module executes
st.set_page_config()/session_state code at import time.
"""

import io
import re
from dataclasses import dataclass
from pathlib import Path
from typing import Dict, List, Optional

import pandas as pd

try:
    from rapidfuzz import fuzz
    HAS_RAPIDFUZZ = True
except ImportError:
    HAS_RAPIDFUZZ = False


# ═══════════════════════════════════════════════════════════════════════════
# Koel Installatie — offerte (bidsheet) vs IB matching
# ═══════════════════════════════════════════════════════════════════════════

@dataclass
class OfferteItem:
    row: int
    categorie: Optional[str]
    subgroep: Optional[str]
    omschrijving: str
    eenheid: Optional[str]
    prijs_materiaal: Optional[float]
    arbeid: Optional[float]
    aantal: Optional[float]
    art_manual: Optional[str]


@dataclass
class IBItem:
    row: int
    art_nr_jumbo: str
    manual_nr: Optional[str]
    art_nr_leverancier: Optional[str]
    groep: Optional[str]
    omschrijving: str
    eenheid: Optional[str]
    aantal: Optional[float]
    materiaalkosten: Optional[float]


def _nh(s) -> str:
    """Normalize a header label for loose matching: lowercase, alnum only."""
    return re.sub(r"[^a-z0-9]+", "", str(s or "").lower())


def _stringify_manual_code(v) -> Optional[str]:
    """Excel sometimes stores an Art.Nr.Jumbo override as a number instead
    of text — e.g. a bare 470 rather than 'Jum-ko-0470', or 19.81 as a real
    float for a buffetten code. Normalize to a string either way so
    downstream matching isn't silently skipped just because of the cell's
    Excel type (a numeric cell would otherwise never equal a string key)."""
    if v is None:
        return None
    if isinstance(v, str):
        return v
    if isinstance(v, float):
        return str(int(v)) if v.is_integer() else str(v)
    return str(v)


_OFFERTE_KEYS = {
    "omschrijving": {"omschrijvingitem"},
    "eenheid": {"eenheid"},
    "prijs_materiaal": {"prijsmateriaalpereenheid"},
    "arbeid": {"arbeidskostenpereenheid"},
    "netto": {"nettonettoprijspereenheidmateriaalenarbeid", "nettonettoprijspereenheidarbeidmateriaal"},
    # 'Aantal in normwinkel' — newer bidsheets (e.g. Carrier's 'Bidsheet K')
    "aantal": {"aantal", "aantalinnormwinkel"},
    "art_manual": {"artnrjumbo", "artjumbo", "artnrjumbomanual"},
}

_IB_KEYS = {
    "art_nr_jumbo": {"artnrjumbo"},
    "manual_nr": {"manualnr"},
    "art_nr_leverancier": {"artnrleverancier"},
    "groep": {"groep"},
    "omschrijving": {"omschrijvingitem"},
    "eenheid": {"eenheid"},
    "aantal": {"aantal"},
    "materiaalkosten": {"materiaalkostenpereenheid"},
}


def _pick_sheet(wb, keyword: str) -> List[str]:
    return [s for s in wb.sheetnames if keyword.lower() in s.lower()]


_KOELING_SHEET_RE = re.compile(r"^\d{4}\s+Installatie\s+\S", re.IGNORECASE)


def _pick_koeling_sheets(wb) -> List[str]:
    """Sheets like '2024 Installatie Frimex' — narrower than a plain
    'Installatie' substring match, which also catches unrelated sheets like
    '2025 B.03 E-installatie SPIE' or '2024 B.01 analyse S-installatie'."""
    strict = [s for s in wb.sheetnames if _KOELING_SHEET_RE.match(s)]
    return strict or _pick_sheet(wb, "Installatie")


def _rows_from_bytes(file_bytes: bytes, sheet_name: str):
    """Yield (1-based row number, row values) via pandas — much faster than
    openpyxl's read_only=True random .cell() access for large/complex
    workbooks (see build_budget_link's docstring for measured numbers)."""
    df = pd.read_excel(io.BytesIO(file_bytes), sheet_name=sheet_name, header=None, engine="openpyxl")
    for i, raw in enumerate(df.values.tolist(), start=1):
        yield i, [None if (isinstance(v, float) and pd.isna(v)) else v for v in raw]


def _build_header_map(normed: List[str], keys: Dict[str, set]) -> Dict[str, int]:
    """Map each logical field name in `keys` to the 1-based column index of
    the first cell in `normed` whose normalized text is one of its variants."""
    header_map: Dict[str, int] = {}
    for key, variants in keys.items():
        for c, n in enumerate(normed, start=1):
            if n in variants:
                header_map[key] = c
                break
    return header_map


def _row_getter(header_map: Dict[str, int], row_vals: list):
    """Return a get(key) closure reading the column `header_map` maps `key`
    to out of `row_vals`, or None if unmapped/out of range."""
    def get(key):
        c = header_map.get(key)
        return row_vals[c - 1] if c and c - 1 < len(row_vals) else None
    return get


#: A Dutch-formatted amount as pdfplumber extracts it from a printed
#: bidsheet: optional '€', thousands '.', decimal ',', and stray spaces
#: inside the digits ('€ 1 2.463,15', '€ 7 ,31', aantal '1 20') — matched
#: after removing all spaces.
_PDF_NUM_RE = re.compile(r"^€?-?[\d.]*\d(,\d+)?$")


def _pdf_cell(v):
    """Turn one PDF table cell into what the Excel reader would give: None
    for empty, a float for amounts/aantallen ('€ -' / '-' = 0), else the
    stripped text."""
    if v is None:
        return None
    s = str(v).strip()
    if not s:
        return None
    compact = s.replace(" ", "")
    if compact in ("-", "€-"):
        return 0.0
    if _PDF_NUM_RE.match(compact):
        return float(compact.replace("€", "").replace(".", "").replace(",", "."))
    return s


def _rows_from_pdf(file_bytes: bytes):
    """Yield (row number, row values) for every table row of a PDF-printed
    bidsheet, all pages chained — a categorie continuing on the next page
    has no header row of its own there, so the parsers keep the previous
    page's column mapping. Row numbers are just a running index."""
    import pdfplumber

    r = 0
    with pdfplumber.open(io.BytesIO(file_bytes)) as pdf:
        for page in pdf.pages:
            for table in page.extract_tables():
                for raw in table:
                    r += 1
                    yield r, [_pdf_cell(v) for v in raw]


def parse_offerte(file_bytes: bytes, sheet_name: str) -> List[OfferteItem]:
    return _parse_offerte_rows(_rows_from_bytes(file_bytes, sheet_name))


def parse_offerte_pdf(file_bytes: bytes) -> List[OfferteItem]:
    """Same as parse_offerte, for a bidsheet printed to PDF (e.g. Frimex
    'Jumbo Hellevoetsluis KT bidsheet 2026.2.pdf')."""
    return _parse_offerte_rows(_rows_from_pdf(file_bytes))


def _parse_offerte_rows(rows) -> List[OfferteItem]:
    items: List[OfferteItem] = []
    category = None
    header_map: Dict[str, int] = {}

    for r, row_vals in rows:
        normed = [_nh(v) for v in row_vals]

        if "omschrijvingitem" in normed:
            header_map = _build_header_map(normed, _OFFERTE_KEYS)
            c1 = row_vals[0]
            if isinstance(c1, str) and c1.strip().lower().startswith("categorie"):
                category = c1.strip()
            continue

        c1 = row_vals[0]
        if isinstance(c1, str) and c1.strip().lower().startswith("categorie"):
            category = c1.strip()
            continue

        # 'Buiten RFP' lines are extra positions outside the tender, handled
        # like NRFP by parse_buiten_rfp — not part of any categorie.
        if _nh(c1) == "buitenrfp":
            continue
        # A PDF-printed bidsheet ends with the NRFP page ('Positions not
        # included in RFP', parsed by _parse_nrfp_rows) — its rows aren't
        # offerte lines, so drop the column mapping from here on.
        if _nh(c1).startswith("positionsnotincludedinrfp"):
            header_map = {}
            continue

        if "omschrijving" not in header_map:
            continue
        desc_col = header_map["omschrijving"]
        desc = row_vals[desc_col - 1] if desc_col - 1 < len(row_vals) else None
        if not desc or not isinstance(desc, str):
            continue

        get = _row_getter(header_map, row_vals)

        art_manual = _stringify_manual_code(get("art_manual"))
        if art_manual:
            art_manual = art_manual.strip() or None
            if art_manual and _nh(art_manual) in _OFFERTE_KEYS["art_manual"]:
                art_manual = None

        prijs_materiaal, arbeid = get("prijs_materiaal"), get("arbeid")
        # Some leveranciers (e.g. Carrier's 'Bidsheet K', Ruimteverdamper
        # vleesafdeling) leave materiaal/arbeid empty and only fill the
        # combined 'Netto netto prijs per eenheid' — use that as the price
        # so the line isn't counted as € 0.
        if prijs_materiaal is None and arbeid is None:
            netto = get("netto")
            if isinstance(netto, (int, float)):
                prijs_materiaal = netto

        items.append(OfferteItem(
            row=r, categorie=category, subgroep=c1 if isinstance(c1, str) else None,
            omschrijving=desc.strip(), eenheid=get("eenheid"),
            prijs_materiaal=prijs_materiaal, arbeid=arbeid,
            aantal=get("aantal"), art_manual=art_manual,
        ))
    return items


def parse_ib(file_bytes: bytes, sheet_name: str) -> List[IBItem]:
    items: List[IBItem] = []
    header_map: Dict[str, int] = {}

    for r, row_vals in _rows_from_bytes(file_bytes, sheet_name):
        normed = [_nh(v) for v in row_vals]

        if "artnrjumbo" in normed and "omschrijvingitem" in normed:
            header_map = _build_header_map(normed, _IB_KEYS)
            continue

        if "art_nr_jumbo" not in header_map or "omschrijving" not in header_map:
            continue

        get = _row_getter(header_map, row_vals)

        art = get("art_nr_jumbo")
        if not isinstance(art, str) or not art.strip():
            continue
        art = art.strip()
        if art.upper().startswith("NRFP"):
            continue

        desc = get("omschrijving")
        if not desc:
            continue

        items.append(IBItem(
            row=r, art_nr_jumbo=art, manual_nr=get("manual_nr"),
            art_nr_leverancier=get("art_nr_leverancier"), groep=get("groep"),
            omschrijving=str(desc).strip(), eenheid=get("eenheid"),
            aantal=get("aantal"), materiaalkosten=get("materiaalkosten"),
        ))
    return items


@dataclass
class NrfpItem:
    code: str
    omschrijving: Optional[str]
    prijs: Optional[float]
    aantal: Optional[float]
    totaal: float


_NRFP_CODE_RE = re.compile(r"^NRFP\s*\d+", re.IGNORECASE)


def parse_nrfp(file_bytes: bytes, sheet_name: str = "NRFP") -> List[NrfpItem]:
    """Parse the 'Positions not included in RFP' sheet — extra/change items
    outside the main tender (huur container, brandsturing, extra led
    verlichting, ...) that the offerte's own grand total still adds in
    (Offertetemplate KT!B12 = '=M16+...+M146+NRFP!E23'). Not part of the
    per-category line items at all, so build_installatie_df never sees
    these — callers that want the true offerte grand total need to add
    sum(item.totaal for item in parse_nrfp(...)) on top separately.
    """
    try:
        rows = list(_rows_from_bytes(file_bytes, sheet_name))
    except ValueError:
        return []
    return _parse_nrfp_rows(rows)


def _parse_nrfp_rows(rows) -> List[NrfpItem]:
    """NRFP rows: code, omschrijving, prijs, aantal, totaal in the first five
    columns — the same layout on the Excel NRFP sheet and the PDF's
    'Positions not included in RFP' page."""
    items: List[NrfpItem] = []
    for _r, row in rows:
        cell = lambda i: row[i] if len(row) > i else None  # noqa: E731
        code = cell(0)
        if not isinstance(code, str) or not _NRFP_CODE_RE.match(code.strip()):
            continue
        totaal = cell(4)
        if not (_is_num(totaal) and totaal):
            continue
        omschrijving = cell(1)
        items.append(NrfpItem(
            code=code.strip(),
            omschrijving=str(omschrijving).strip() if isinstance(omschrijving, str) else None,
            prijs=cell(2) if _is_num(cell(2)) else None,
            aantal=cell(3) if _is_num(cell(3)) else None,
            totaal=float(totaal),
        ))
    return items


def parse_buiten_rfp(file_bytes: bytes, sheet_name: str) -> List[NrfpItem]:
    """Parse 'Buiten RFP' lines on the bidsheet itself (e.g. Carrier's
    'Bidsheet K' row 140 'Afblazen bestaande installatie') — the same kind of
    out-of-tender position as the NRFP sheet, so returned as NrfpItem for
    callers to add alongside parse_nrfp's items.

    The 'Buiten RFP' header row only repeats a few column labels (no
    Aantal/Prijs), so each header row is merged into the mapping from the
    earlier full categorie headers instead of replacing it — the columns sit
    in the same position across the sheet."""
    return _parse_buiten_rfp_rows(_rows_from_bytes(file_bytes, sheet_name))


def _parse_buiten_rfp_rows(rows) -> List[NrfpItem]:
    items: List[NrfpItem] = []
    header_map: Dict[str, int] = {}

    for _r, row_vals in rows:
        normed = [_nh(v) for v in row_vals]
        if "omschrijvingitem" in normed:
            header_map.update(_build_header_map(normed, _OFFERTE_KEYS))
            continue
        if not row_vals or _nh(row_vals[0]) != "buitenrfp":
            continue

        get = _row_getter(header_map, row_vals)
        desc = get("omschrijving")
        if not isinstance(desc, str) or not desc.strip():
            continue
        materiaal, arbeid, netto = get("prijs_materiaal"), get("arbeid"), get("netto")
        num = lambda x: x if isinstance(x, (int, float)) else None  # noqa: E731
        if num(materiaal) is not None or num(arbeid) is not None:
            prijs = (num(materiaal) or 0) + (num(arbeid) or 0)
        else:
            prijs = num(netto)
        aantal = num(get("aantal"))
        if not prijs or not aantal:
            continue
        items.append(NrfpItem(
            code="Buiten RFP", omschrijving=desc.strip(),
            prijs=prijs, aantal=aantal, totaal=float(prijs * aantal),
        ))
    return items


def is_pdf(filename: str) -> bool:
    return filename.lower().endswith(".pdf")


def read_bidsheet(file_bytes: bytes, filename: str, sheet_name: Optional[str] = None) -> tuple:
    """Read a Koel Installatie bidsheet, Excel or PDF, into (offerte items,
    NRFP + Buiten RFP items). For Excel, `sheet_name` defaults to the first
    pick_bid_sheet candidate; a PDF has no sheets (all pages are read)."""
    if is_pdf(filename):
        rows = list(_rows_from_pdf(file_bytes))
        return _parse_offerte_rows(rows), _parse_nrfp_rows(rows) + _parse_buiten_rfp_rows(rows)
    if sheet_name is None:
        import openpyxl

        wb = openpyxl.load_workbook(io.BytesIO(file_bytes), read_only=True)
        sheet_name = pick_bid_sheet(wb.sheetnames)[0]
    items = parse_offerte(file_bytes, sheet_name)
    return items, parse_nrfp(file_bytes) + parse_buiten_rfp(file_bytes, sheet_name)


@dataclass
class BudgetLink:
    budget_row: int
    code: Optional[str]
    omschrijving: Optional[str]
    aantal: Optional[float]
    prijs: Optional[float] = None   # Budget!L — leverancier-resolved prijs per eenheid
    totaal: Optional[float] = None  # Budget!M — Budget's own ROUND(L,2)*ROUND(K,2)


def _budget_prijs_totaal(row) -> tuple:
    """Budget!L (col 12, Prijs) and Budget!M (col 13, Totaal) for a row
    already fetched via _budget_row — this is the IB-side comparison figure
    the Samenvatting should use, not a recomputation from the installatie
    sheet (which is materiaal-only and ignores the Budget's own leverancier/
    aantal resolution)."""
    prijs = row.iloc[11] if len(row) > 11 else None
    totaal = row.iloc[12] if len(row) > 12 else None
    return (prijs if pd.notna(prijs) else None, totaal if pd.notna(totaal) else None)


#: Budget-sheet rows covering the Koeling installatie section (C.03 Koel- en
#: vriescellen through C.05 Koelinstallatie) — same convention as Van Keulen's
#: hardcoded Budget row range for its own section.
BUDGET_KOELING_ROW_START = 1513
BUDGET_KOELING_ROW_END = 1716


_LEADING_CODE_RE = re.compile(r"^(\d{3,4})\b")
_ART_SUFFIX_RE = re.compile(r"(\d{3,4})$")


def load_budget_df(ib_file_bytes: bytes) -> Optional[pd.DataFrame]:
    """Load the Budget sheet once via pandas — see build_budget_link's
    docstring for why (openpyxl read_only random-cell access on this sheet
    takes minutes; pd.read_excel takes ~1s). Shared by build_budget_link and
    build_buffetten_link so a single analysis run only pays this cost once."""
    try:
        return pd.read_excel(io.BytesIO(ib_file_bytes), sheet_name="Budget", header=None, engine="openpyxl")
    except ValueError:
        return None


def _budget_row(bdf: pd.DataFrame, r: int) -> Optional[tuple]:
    """Return (desc, aantal, row) for 1-based Budget row `r` — desc is a
    stripped non-blank string (column H), aantal is column K's raw value
    (or None if NaN), row is the full pandas row for any other columns a
    caller needs. None if `r` is out of bounds or H is blank/non-string."""
    idx = r - 1
    if idx >= len(bdf):
        return None
    row = bdf.iloc[idx]
    desc = row.iloc[7] if len(row) > 7 else None
    if not isinstance(desc, str) or not desc.strip():
        return None
    aantal = row.iloc[10] if len(row) > 10 else None
    return desc.strip(), (aantal if pd.notna(aantal) else None), row


def _budget_code_totals(bdf: pd.DataFrame, row_start: int, row_end: int) -> Dict[str, tuple]:
    """Map every Budget row's own code (col G, e.g. 'C.04.04.01') to
    (desc, aantal, totaal), across the whole range including group-header
    rows whose description has no leading article number (e.g. 'Gascooler',
    code 'C.04.04') — those are excluded from the leaf-level scan in
    build_budget_link but are exactly what _budget_group_totaal needs."""
    totals: Dict[str, tuple] = {}
    for r in range(row_start, row_end + 1):
        idx = r - 1
        if idx >= len(bdf):
            break
        row = bdf.iloc[idx]
        code = row.iloc[6] if len(row) > 6 else None
        if not isinstance(code, str) or not code.strip():
            continue
        desc = row.iloc[7] if len(row) > 7 else None
        aantal = row.iloc[10] if len(row) > 10 else None
        _, totaal = _budget_prijs_totaal(row)
        totals[code.strip()] = (
            desc.strip() if isinstance(desc, str) else None,
            aantal if pd.notna(aantal) else None,
            totaal,
        )
    return totals




def _find_alt_budget_row(
    bdf: pd.DataFrame, row_start: int, row_end: int, desc: Optional[str],
    used_rows: set, threshold: int = 55,
) -> Optional[BudgetLink]:
    """Fail-safe for build_budget_link: the numeric-suffix join can land on a
    Budget row IB didn't actually select (e.g. offerte quotes 'Bypass valve
    5' but IB's own aantal is on the 'Bypass valve 4' row instead — a
    different, similarly-described variant), which shows up as that link's
    aantal being blank/0. Fall back to the closest text match, among Budget
    rows not already claimed by another link, that does carry a real aantal."""
    if not desc:
        return None
    target = _norm_text(desc)
    best, best_score = None, threshold
    for r in range(row_start, row_end + 1):
        if r in used_rows:
            continue
        result = _budget_row(bdf, r)
        if result is None:
            continue
        cdesc, aantal, row = result
        if not aantal:
            continue
        score = fuzz.token_set_ratio(target, _norm_text(cdesc))
        if score > best_score:
            prijs, totaal = _budget_prijs_totaal(row)
            best_score = score
            best = BudgetLink(
                budget_row=r, code=row.iloc[6] if len(row) > 6 else None,
                omschrijving=cdesc, aantal=aantal, prijs=prijs, totaal=totaal,
            )
    return best


def build_budget_link(
    bdf: Optional[pd.DataFrame], ib_items: List[IBItem],
    row_start: int = BUDGET_KOELING_ROW_START, row_end: int = BUDGET_KOELING_ROW_END,
) -> Dict[int, BudgetLink]:
    """Map installatie-sheet row -> the Budget-sheet row (1513-1716, the
    Koeling section) that prices it.

    Budget!H (omschrijving) starts each line with the same 4-digit number as
    the last 4 digits of the item's Art.Nr.Jumbo — e.g. Budget "0470 Gas
    Cooler Ec fans 133kW" <-> installatie row 'Jum-ko-0470'. That numeric
    prefix is the join key; no formula-parsing needed.
    """
    if bdf is None:
        return {}

    suffix_to_row: Dict[str, int] = {}
    for item in ib_items:
        m = _ART_SUFFIX_RE.search(item.art_nr_jumbo)
        if m:
            suffix_to_row[m.group(1).zfill(4)] = item.row
    ib_by_row = {item.row: item for item in ib_items}

    links: Dict[int, BudgetLink] = {}
    used_budget_rows: set = set()
    for r in range(row_start, row_end + 1):
        result = _budget_row(bdf, r)
        if result is None:
            continue
        desc, aantal, row = result
        m = _LEADING_CODE_RE.match(desc)
        if not m:
            continue
        ib_row = suffix_to_row.get(m.group(1).zfill(4))
        if ib_row is None:
            continue
        prijs, totaal = _budget_prijs_totaal(row)
        links[ib_row] = BudgetLink(
            budget_row=r, code=row.iloc[6] if len(row) > 6 else None,
            omschrijving=desc, aantal=aantal, prijs=prijs, totaal=totaal,
        )
        used_budget_rows.add(r)

    if HAS_RAPIDFUZZ:
        for ib_row, link in list(links.items()):
            if link.aantal:
                continue
            ib_item = ib_by_row.get(ib_row)
            alt = _find_alt_budget_row(
                bdf, row_start, row_end,
                ib_item.omschrijving if ib_item else link.omschrijving,
                used_budget_rows,
            )
            if alt is not None:
                used_budget_rows.discard(link.budget_row)
                used_budget_rows.add(alt.budget_row)
                links[ib_row] = alt

    return links


#: Koelbuffetten (Categorie buffetten) live outside the installatie sheet
#: entirely — Frimex quotes furniture separately from koelinstallatie — so
#: there's no Art.Nr.Jumbo to join on. These two Budget-sheet ranges price
#: them directly.
BUDGET_BUFFETTEN_RANGES = [(1442, 1455), (1503, 1512)]

_DEURS_RE = re.compile(r"(\d+)-deurs", re.IGNORECASE)
_VARIANT_DRS_RE = re.compile(r"variant\s*(\d+)\s*drs", re.IGNORECASE)


_BUDGET_LEAD_CODE_RE = re.compile(r"^(\d+(?:\.\d+)?)")
_BUFFET_MANUAL_CODE_RE = re.compile(r"^(\d+(?:\.\d+)?)(?:-(\d+)drs)?$", re.IGNORECASE)


def _norm_code(raw) -> Optional[str]:
    """'19.8', '19.80', 19.81 (float) -> '19.80'/'19.81' — Budget's leading
    code is always written with 2 decimals, so normalize through float."""
    try:
        return f"{float(raw):.2f}"
    except (TypeError, ValueError):
        return None


def _parse_buffet_manual_code(art_manual) -> Optional[tuple]:
    """Parse a manually-entered buffetten override from Art.Nr.Jumbo:
    '19.81' / 19.81 (Excel auto-numbered) -> ('19.81', None); '19.80-4drs'
    -> ('19.80', '4'). Returns None if it doesn't look like a buffetten code
    (e.g. it's a Jum-ko-XXXX installatie code, handled elsewhere)."""
    if art_manual is None:
        return None
    if isinstance(art_manual, (int, float)):
        return (_norm_code(art_manual), None)
    m = _BUFFET_MANUAL_CODE_RE.match(str(art_manual).strip())
    if not m:
        return None
    code_raw, drs = m.groups()
    return (_norm_code(code_raw), drs)


def build_buffetten_link(bdf: Optional[pd.DataFrame], offerte: List[OfferteItem]) -> Dict[int, tuple]:
    """Match 'Categorie buffetten' offerte lines to Budget rows 1442-1455 /
    1503-1512, keyed by offerte row (there's no installatie-sheet item to
    bridge through, unlike build_budget_link). No Jum-ko-style Art.Nr.Jumbo
    exists for these, but the same column can carry a Budget product code
    instead (e.g. '19.81', or '19.80-4drs' when a code has door-count
    variants) — that's matched first, as an unambiguous manual override.
    Items without one fall back to rule-based text matching:
      - "Koelbuffet <N>-deurs" <-> Budget "...Prijs variant <N> drs" (same
        door count)
      - "...Saladiere..." <-> Budget "...saladiere..." (substring)

    Returns {offerte_row: (BudgetLink, match_source)} where match_source is
    "handmatig" (via the manual code) or "budget-tekst" (heuristic).
    """
    if bdf is None:
        return {}

    entries = []
    for r_start, r_end in BUDGET_BUFFETTEN_RANGES:
        for r in range(r_start, r_end + 1):
            result = _budget_row(bdf, r)
            if result is None:
                continue
            desc, aantal, row = result
            cm = _BUDGET_LEAD_CODE_RE.match(desc)
            vm = _VARIANT_DRS_RE.search(desc)
            prijs, totaal = _budget_prijs_totaal(row)
            entries.append({
                "row": r, "desc": desc, "aantal": aantal, "prijs": prijs, "totaal": totaal,
                "code": _norm_code(cm.group(1)) if cm else None,
                "variant": vm.group(1) if vm else None,
            })

    used_budget = set()
    links: Dict[int, tuple] = {}

    def _claim(entry, off_row, source):
        used_budget.add(entry["row"])
        link = BudgetLink(
            budget_row=entry["row"], code=entry["code"], omschrijving=entry["desc"],
            aantal=entry["aantal"], prijs=entry["prijs"], totaal=entry["totaal"],
        )
        links[off_row] = (link, source)

    buffetten = [o for o in offerte if o.categorie and "buffet" in o.categorie.lower()]

    # Pass 1: manual code override (unambiguous — takes priority).
    for o in buffetten:
        parsed = _parse_buffet_manual_code(o.art_manual)
        if not parsed:
            continue
        code, drs = parsed
        candidates = [
            e for e in entries
            if e["row"] not in used_budget and e["code"] == code and (drs is None or e["variant"] == drs)
        ]
        if len(candidates) == 1:
            _claim(candidates[0], o.row, "handmatig")

    # Pass 2: heuristic text fallback for whatever's left.
    for o in buffetten:
        if o.row in links:
            continue
        desc_l = o.omschrijving.lower()

        m = _DEURS_RE.search(desc_l)
        if m and "koelbuffet" in desc_l:
            n = m.group(1)
            for e in entries:
                if e["row"] not in used_budget and e["variant"] == n and "koelbuffet" in e["desc"].lower():
                    _claim(e, o.row, "budget-tekst")
                    break
            continue

        if "saladiere" in desc_l:
            for e in entries:
                if e["row"] not in used_budget and "saladiere" in e["desc"].lower():
                    _claim(e, o.row, "budget-tekst")
                    break

    return links


#: Budget-sheet subheading rows (the C.0X.0Y "group" level, e.g. 'C.04.04
#: Gascooler') span this range across both the koelbuffetten and koeling
#: sections — see load_categorie_budget_codes.
CATEGORIE_BUDGET_ROW_START = 1442
CATEGORIE_BUDGET_ROW_END = 1716


def load_categorie_budget_codes() -> Dict[str, List[tuple]]:
    """Read data/koeling_categorie_budget_codes.csv: offerte Categorie label
    -> [(budget_code, budget_omschrijving), ...]. Each budget_code is a
    Budget!G outline code (e.g. 'C.04.04') for a subheading row whose own
    Totaal (col M) already sums whichever child variant was actually
    selected — the figure that should represent that category's IB total,
    not a per-offerte-row rollup that can read 0 when IB chose a different
    variant than the offerte quoted (see build_categorie_totalen).

    Keyed by the Budget outline *code*, not a row number, so this keeps
    working if rows shift when the workbook is edited — build_categorie_totalen
    looks the code up fresh in the current sheet each run. budget_omschrijving
    is kept as a documentation/fallback aid: if a code is ever renumbered,
    matching by its still-recognizable description keeps the lookup working
    without needing to update this file."""
    path = ART_NR_MAPPING_DIR / "koeling_categorie_budget_codes.csv"
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


def build_categorie_totalen(
    bdf: Optional[pd.DataFrame], mapping: Dict[str, List[tuple]],
    row_start: int = CATEGORIE_BUDGET_ROW_START, row_end: int = CATEGORIE_BUDGET_ROW_END,
) -> Dict[str, float]:
    """Sum each offerte Categorie's own Budget subheading Totaal(s) directly
    — e.g. 'Categorie gascoolers' -> Budget!C.04.04's own Totaal (26,564.64),
    which already reflects whichever variant was actually selected — instead
    of the per-offerte-row rollup in build_installatie_df, which reads 0 for
    a specific variant the offerte quoted but IB ended up not choosing.
    Categories with no entry in `mapping` are simply absent from the result,
    so callers should fall back to the per-row rollup for those."""
    if bdf is None or not mapping:
        return {}
    code_totals = _budget_code_totals(bdf, row_start, row_end)
    desc_totals: Dict[str, float] = {}
    for desc, _aantal, totaal in code_totals.values():
        if desc and totaal:
            desc_totals[desc] = totaal

    result: Dict[str, float] = {}
    for categorie, entries in mapping.items():
        total = 0.0
        found = False
        for code, omschrijving in entries:
            entry = code_totals.get(code)
            totaal = entry[2] if entry else None
            if not totaal and omschrijving:
                totaal = desc_totals.get(omschrijving)
            if totaal:
                total += totaal
                found = True
        if found:
            result[categorie] = total
    return result


def _norm_text(s) -> str:
    if s is None:
        return ""
    s = str(s).replace("₂", "2")
    s = re.sub(r"[^\w\s]", " ", s, flags=re.UNICODE)
    return re.sub(r"\s+", " ", s).strip().lower()


#: Persistent Omschrijving -> Art.Nr. lookup per leverancier, hand-maintained
#: as matches get manually confirmed over time. Used as a fallback when the
#: uploaded offerte has no Art.Nr.Jumbo column filled in — the normal case,
#: since the original bidsheet as delivered by the leverancier doesn't have
#: that column at all; it only exists in copies we've annotated by hand.
ART_NR_MAPPING_DIR = Path(__file__).parent / "data"


def load_art_nr_mapping(leverancier: str = "frimex") -> Dict[str, str]:
    path = ART_NR_MAPPING_DIR / f"koeling_art_nr_mapping_{leverancier}.csv"
    if not path.exists():
        return {}
    df = pd.read_csv(path)
    return {_norm_text(row["omschrijving"]): str(row["art_nr"]).strip() for _, row in df.iterrows()}


def apply_art_nr_mapping(offerte: List[OfferteItem], mapping: Dict[str, str]) -> int:
    """Fill in art_manual from the external mapping for rows that don't
    already have one set from the offerte file itself (which takes priority
    — a project-specific override always wins over the shared dictionary).
    Returns how many rows got filled in this way."""
    if not mapping:
        return 0
    filled = 0
    for o in offerte:
        if o.art_manual:
            continue
        code = mapping.get(_norm_text(o.omschrijving))
        if code:
            o.art_manual = code
            filled += 1
    return filled


def _numeric_tokens(s: str) -> set:
    return set(re.findall(r"\d+", s))


def _score_pair(off: OfferteItem, ib: IBItem) -> float:
    a = _norm_text(off.subgroep) + " " + _norm_text(off.omschrijving)
    b = _norm_text(ib.groep) + " " + _norm_text(ib.omschrijving)
    text_score = fuzz.token_set_ratio(a, b)

    na, nb = _numeric_tokens(a), _numeric_tokens(b)
    if na or nb:
        num_score = 100 * len(na & nb) / (len(na | nb) or 1)
    else:
        num_score = 50

    unit_bonus = 0
    ua, ub = _norm_text(off.eenheid), _norm_text(ib.eenheid)
    if ua and ub and ua[:1] == ub[:1]:
        unit_bonus = 5

    return 0.55 * text_score + 0.40 * num_score + unit_bonus


_JUM_KO_RE = re.compile(r"^jum-ko-(\d+)$", re.IGNORECASE)
_BARE_DIGITS_RE = re.compile(r"^(\d+)$")


def _normalize_jum_ko_code(raw: str) -> Optional[str]:
    """'Jum-ko-0470' (any case/padding) -> 'Jum-ko-0470'; a bare number like
    '470' -> 'Jum-ko-0470' too, so a manual override typed (or auto-typed by
    Excel) as just a number still resolves to the right installatie item.
    None if it's neither shape (e.g. a buffetten code like '19.81')."""
    s = raw.strip()
    m = _JUM_KO_RE.match(s) or _BARE_DIGITS_RE.match(s)
    return f"Jum-ko-{int(m.group(1)):04d}" if m else None


def match_koeling_installatie(offerte: List[OfferteItem], ib: List[IBItem], fuzzy_threshold: int = 40):
    """Manual Art.Nr.Jumbo exact join first (ground truth), then fuzzy
    description fallback for whatever's left. Returns (matches, used_ib_idx)
    where matches: {offerte_idx: (ib_idx, score, source)}."""
    ib_by_art: Dict[str, int] = {}
    for j, item in enumerate(ib):
        ib_by_art.setdefault(item.art_nr_jumbo, j)

    used_off, used_ib = set(), set()
    matches: Dict[int, tuple] = {}

    for i, o in enumerate(offerte):
        if not o.art_manual:
            continue
        j = ib_by_art.get(o.art_manual)
        if j is None:
            normalized = _normalize_jum_ko_code(o.art_manual)
            if normalized:
                j = ib_by_art.get(normalized)
        if j is None:
            continue
        used_off.add(i); used_ib.add(j)
        matches[i] = (j, 100.0, "handmatig")

    if HAS_RAPIDFUZZ:
        candidates = []
        for i, o in enumerate(offerte):
            if i in used_off:
                continue
            for j, item in enumerate(ib):
                if j in used_ib:
                    continue
                sc = _score_pair(o, item)
                if sc >= fuzzy_threshold:
                    candidates.append((sc, i, j))
        candidates.sort(key=lambda x: -x[0])
        for sc, i, j in candidates:
            if i in used_off or j in used_ib:
                continue
            used_off.add(i); used_ib.add(j)
            matches[i] = (j, sc, "fuzzy")

    return matches, used_ib


def build_installatie_df(
    offerte: List[OfferteItem], ib: List[IBItem], matches: Dict[int, tuple],
    budget_link: Optional[Dict[int, BudgetLink]] = None,
    buffetten_link: Optional[Dict[int, tuple]] = None,
) -> pd.DataFrame:
    budget_link = budget_link or {}
    buffetten_link = buffetten_link or {}
    rows = []
    for i, o in enumerate(offerte):
        m = matches.get(i)
        ib_item = ib[m[0]] if m else None
        link = budget_link.get(ib_item.row) if ib_item else None
        # Buffetten items never match an installatie-sheet row (they're not
        # in that sheet at all), so only look here when there's no ib_item.
        buf_entry = buffetten_link.get(o.row) if ib_item is None else None
        buf_link, buf_source = buf_entry if buf_entry else (None, None)

        aantal_o = o.aantal or 0
        active_link = link or buf_link
        aantal_i = active_link.aantal if (active_link and active_link.aantal is not None) else None
        aantal_diff = (aantal_o - aantal_i) if aantal_i is not None else None

        # IB omschrijving comes from the Budget sheet (the actual priced line
        # item as budgeted for this project); fall back to the installatie
        # sheet's own description when no Budget row references it.
        ib_omschrijving = active_link.omschrijving if active_link else (ib_item.omschrijving if ib_item else None)

        prijs_o = o.prijs_materiaal
        # IB Prijs p.e. comes from the Budget sheet's own leverancier-
        # resolved price (column L) — not recomputed from the installatie
        # sheet's materiaal-only column, which ignores labor entirely and
        # isn't what the Budget tab itself reports. Falls back to the
        # installatie sheet's materiaalkosten only when no Budget row
        # references this item at all.
        if active_link and active_link.prijs is not None:
            prijs_i = active_link.prijs
        else:
            prijs_i = ib_item.materiaalkosten if ib_item else None
        prijs_diff = (prijs_o - prijs_i) if (prijs_o is not None and prijs_i is not None) else None
        prijs_diff_pct = (prijs_diff / prijs_i * 100) if (prijs_diff is not None and prijs_i) else None

        # A tiny aantal difference (e.g. 1 extra meter of pipe on a 111m
        # run) isn't worth flagging — only call it out past a 10% deviation
        # from the IB/Budget quantity. When that quantity is 0 (no base to
        # take a percentage of), any nonzero offerte aantal is significant.
        if aantal_diff is None or aantal_diff == 0:
            significant_aantal_diff = False
        elif aantal_i:
            significant_aantal_diff = abs(aantal_diff) / aantal_i * 100 > 10
        else:
            significant_aantal_diff = True

        # buf_link (buffetten) rows now carry a real Budget-sourced prijs
        # just like ib_item rows do, so the same price/aantal checks apply
        # uniformly — only a genuinely unmatched row (neither) skips them.
        if ib_item is None and buf_link is None:
            status = "unmatched"
        elif prijs_diff is not None and abs(prijs_diff) > 0.01:
            status = "prijs_afwijking"
        elif significant_aantal_diff:
            status = "aantal_afwijking"
        else:
            status = "ok"

        # Matches the offerte's own "Totaalprijs winkel in €" column, which
        # is (Prijs materiaal + Arbeidskosten) per eenheid × Aantal — not
        # materiaal alone.
        offerte_totaal = ((prijs_o or 0) + (o.arbeid or 0)) * aantal_o if aantal_o else None
        # Budget's own Totaal (column M = ROUND(L,2)*ROUND(K,2), using
        # Budget's own aantal) — the actual figure booked in the Budget tab
        # for this line, not a recomputation with the offerte's aantal. When
        # the offerte's quoted variant isn't the one IB actually chose (this
        # leaf's own Totaal is 0), this row-level figure understates the
        # category — the page's "Totaal per categorie" table corrects for
        # that separately via build_categorie_totalen, which reads the
        # Budget subheading's own Totaal instead of rolling up leaf rows.
        if active_link and active_link.totaal is not None:
            ib_totaal = active_link.totaal
        else:
            ib_totaal = (prijs_i or 0) * aantal_o if (aantal_o and prijs_i is not None) else None

        if m:
            match_label, match_pct = m[2], f"{m[1]:.0f}%"
        elif buf_link:
            match_label, match_pct = buf_source, ""
        else:
            match_label, match_pct = "geen match", ""

        rows.append({
            "_status": status,
            "_offerte_totaal": offerte_totaal,
            "_ib_totaal": ib_totaal,
            "Categorie": o.categorie,
            "Offerte omschrijving": o.omschrijving,
            "IB omschrijving": ib_omschrijving,
            "Art.Nr.Jumbo": ib_item.art_nr_jumbo if ib_item else None,
            "Match": match_label,
            "Match %": match_pct,
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


# ═══════════════════════════════════════════════════════════════════════════
# Offertevergelijking — meerdere bidsheets naast elkaar vs IB
# ═══════════════════════════════════════════════════════════════════════════

#: Koelcellen, vriescellen and celaccessoires are one physical scope from
#: the koelinstallateur's perspective (the cellen plus the accessories that
#: go with them) even though the offerte/Budget book them as three separate
#: categories — grouped together in every per-categorie overview so they
#: read as one row/slice instead of three.
CATEGORIE_GROEP_LABEL = "Categorie koel-/vriescellen (incl. celaccessoires)"
CATEGORIE_GROEP = {
    "Categorie koelcellen": CATEGORIE_GROEP_LABEL,
    "Categorie vriescellen": CATEGORIE_GROEP_LABEL,
    "Categorie celaccessoires": CATEGORIE_GROEP_LABEL,
}

#: Leveranciers recognised in a bidsheet filename, for a default offerte name.
_LEVERANCIER_HINTS = {"frimex": "Frimex", "carrier": "Carrier", "ccr": "Carrier", "kaplanlaar": "Kaplanlaar"}


def pick_bid_sheet(sheetnames: List[str]) -> List[str]:
    """Candidate offerte sheets in a bidsheet: 'Offertetemplate KT' when
    present, else every sheet (e.g. Carrier's single 'Bidsheet K')."""
    return [s for s in sheetnames if "offertetemplate" in s.lower()] or list(sheetnames)


def _hint_leverancier(text: str) -> Optional[str]:
    low = text.lower()
    for hint, naam in _LEVERANCIER_HINTS.items():
        if re.search(rf"(^|[^a-z]){hint}([^a-z]|$)", low):
            return naam
    return None


def guess_leverancier(filename: str) -> str:
    """Default offerte name from a bidsheet filename — a known leverancier
    when the name mentions one, else the filename without extension."""
    return _hint_leverancier(filename) or re.sub(r"\.(xlsx?|pdf)$", "", filename, flags=re.IGNORECASE)[:30]


def bidsheet_leverancier(file_bytes: bytes, filename: str) -> str:
    """Offerte name from the bidsheet's own 'Leverancier' line near the top
    (e.g. 'Frimex Installatie B.V.' -> 'Frimex', 'Carrier Commercial
    Refrigeration' -> 'Carrier'), falling back to guess_leverancier — the
    Hellevoetsluis PDF's filename doesn't mention the leverancier at all."""
    try:
        if is_pdf(filename):
            rows = _rows_from_pdf(file_bytes)
        else:
            import openpyxl

            sheet = pick_bid_sheet(openpyxl.load_workbook(io.BytesIO(file_bytes), read_only=True).sheetnames)[0]
            rows = _rows_from_bytes(file_bytes, sheet)
        for r, row in rows:
            if r > 40:
                break
            if len(row) > 1 and _nh(row[0]) == "leverancier" and isinstance(row[1], str) and row[1].strip():
                naam = row[1].strip()
                return _hint_leverancier(naam) or naam[:30]
    except Exception:  # a name is a convenience — never fail the export over it
        pass
    return guess_leverancier(filename)


@dataclass
class VergelijkingRegel:
    """One line of the side-by-side comparison: an IB article (or, without
    an IB match, an offerte omschrijving) with each offerte's aantal and
    prijs p.e. — `offertes[i]` is (aantal, prijs) for offerte i, (None, None)
    when that offerte doesn't have the line at all."""
    categorie: Optional[str]
    art_nr_jumbo: Optional[str]
    omschrijving: str
    ib_omschrijving: Optional[str]
    eenheid: Optional[str]
    ib_aantal: Optional[float]
    ib_prijs: Optional[float]
    offertes: List[tuple]
    opmerking: str = ""

    @property
    def ib_totaal(self) -> float:
        return (self.ib_aantal or 0) * (self.ib_prijs or 0)

    def offerte_totaal(self, i: int) -> float:
        aantal, prijs = self.offertes[i]
        return (aantal or 0) * (prijs or 0)


@dataclass
class OfferteVergelijking:
    namen: List[str]
    bestanden: List[str]
    regels: List[VergelijkingRegel]
    nrfp: List[List[NrfpItem]]
    #: grouped Categorie labels in offerte order
    categorieen: List[str]
    #: grouped Categorie -> Budget subheading Totaal, only where every member
    #: categorie has a Budget code (see build_categorie_totalen)
    ib_categorie_totalen: Dict[str, float]
    ib_categorie_codes: Dict[str, str]
    ib_bestand: str = ""


def _is_num(v) -> bool:
    return isinstance(v, (int, float)) and not pd.isna(v)


def _clean(v):
    return None if isinstance(v, float) and pd.isna(v) else v


def build_offerte_vergelijking(
    offertes: List[tuple],
    categorie_totalen: Dict[str, float],
    categorie_mapping: Dict[str, List[tuple]],
    ib_bestand: str = "",
) -> OfferteVergelijking:
    """Line up several bidsheets against the same IB.

    `offertes` is a list of (naam, bestand, items, df, nrfp_items) where
    `df` is build_installatie_df's output for `items` (row-aligned). Lines
    are keyed by the IB Art.Nr.Jumbo each offerte line matched — so the same
    IB article lines up across offertes even if the leveranciers word it
    differently — falling back to the offerte omschrijving for lines without
    an IB match. Such an omschrijving-only line is folded into the IB-article
    line of another offerte with the same omschrijving (they share the
    bidsheet template). Only lines with an aantal in at least one offerte or
    in the IB are kept."""
    per_offerte = []
    for _naam, _bestand, items, df, _nrfp in offertes:
        recs: Dict[str, dict] = {}
        for o, (_, r) in zip(items, df.iterrows(), strict=True):
            art = r["Art.Nr.Jumbo"] if isinstance(r["Art.Nr.Jumbo"], str) else None
            key = art or "omschr:" + _nh(o.omschrijving)
            prijs = None
            if o.prijs_materiaal is not None or o.arbeid is not None:
                prijs = (o.prijs_materiaal or 0) + (o.arbeid or 0)
            aantal = o.aantal or None
            rec = recs.get(key)
            if rec is None:
                recs[key] = dict(
                    art=art, categorie=CATEGORIE_GROEP.get(o.categorie, o.categorie),
                    omschrijving=o.omschrijving, eenheid=o.eenheid, aantal=aantal, prijs=prijs, row=o.row,
                    ib_omschrijving=_clean(r["IB omschrijving"]), ib_aantal=_clean(r["IB Aantal"]),
                    ib_prijs=_clean(r["IB Prijs p.e."]), match=r["Match"], match_pct=r["Match %"], n=1,
                )
            elif aantal:
                # Two offerte lines on one IB article: add them up.
                if rec["aantal"]:
                    totaal = rec["aantal"] * (rec["prijs"] or 0) + aantal * (prijs or 0)
                    rec["aantal"] += aantal
                    rec["prijs"] = totaal / rec["aantal"]
                else:
                    rec.update(aantal=aantal, prijs=prijs)
                rec["n"] += 1
        per_offerte.append(recs)

    by_desc: Dict[str, str] = {}
    for recs in per_offerte:
        for k, rec in recs.items():
            if not k.startswith("omschr:"):
                by_desc.setdefault(_nh(rec["omschrijving"]), k)
    slots: Dict[str, Dict[int, dict]] = {}
    for i, recs in enumerate(per_offerte):
        for k, rec in recs.items():
            target = by_desc.get(k[len("omschr:"):]) if k.startswith("omschr:") else None
            if target and target not in recs:
                k = target
            slots.setdefault(k, {})[i] = rec

    def first(slot, field):
        for i in sorted(slot):
            v = slot[i].get(field)
            if v is not None:
                return v
        return None

    namen = [o[0] for o in offertes]
    regels, order = [], []
    for slot in slots.values():
        ib_aantal = first(slot, "ib_aantal")
        if not any(rec["aantal"] for rec in slot.values()) and not ib_aantal:
            continue
        notes = []
        for i in sorted(slot):
            rec = slot[i]
            if rec["match"] == "fuzzy":
                notes.append(f"{namen[i]}: fuzzy match {rec['match_pct']}")
            elif rec["match"] == "geen match" and rec["aantal"]:
                notes.append(f"{namen[i]}: geen IB-match")
            if rec["n"] > 1:
                notes.append(f"{namen[i]}: {rec['n']} regels samengeteld")
        regels.append(VergelijkingRegel(
            categorie=first(slot, "categorie"), art_nr_jumbo=first(slot, "art"),
            omschrijving=first(slot, "omschrijving"), ib_omschrijving=first(slot, "ib_omschrijving"),
            eenheid=first(slot, "eenheid"), ib_aantal=ib_aantal, ib_prijs=first(slot, "ib_prijs"),
            offertes=[(slot[i]["aantal"], slot[i]["prijs"]) if i in slot else (None, None)
                      for i in range(len(offertes))],
            opmerking="; ".join(notes),
        ))
        i0 = min(slot)
        order.append((i0, slot[i0]["row"]))

    # Keep bidsheet order: categories as they first appear, lines by row.
    ranked = sorted(zip(order, regels, strict=True), key=lambda x: x[0])
    categorieen: List[str] = []
    for _, regel in ranked:
        if regel.categorie not in categorieen:
            categorieen.append(regel.categorie)
    ranked.sort(key=lambda x: (categorieen.index(x[1].categorie), x[0]))

    ib_tot, ib_codes = {}, {}
    for cat in categorieen:
        members = [k for k, g in CATEGORIE_GROEP.items() if g == cat] or [cat]
        if all(m in categorie_totalen for m in members):
            ib_tot[cat] = round(sum(categorie_totalen[m] for m in members), 2)
            ib_codes[cat] = ", ".join(code for m in members for code, _ in categorie_mapping.get(m, []))

    return OfferteVergelijking(
        namen=namen, bestanden=[o[1] for o in offertes], regels=[r for _, r in ranked],
        nrfp=[o[4] for o in offertes], categorieen=categorieen,
        ib_categorie_totalen=ib_tot, ib_categorie_codes=ib_codes, ib_bestand=ib_bestand,
    )


def _laagste(namen: List[str], totalen: List[float]) -> str:
    quoted = [(t, n) for t, n in zip(totalen, namen, strict=True) if t > 0]
    if not quoted:
        return ""
    best = min(quoted, key=lambda x: x[0])[1]
    return f"Alleen {best}" if len(quoted) == 1 else best


def vergelijking_samenvatting(v: OfferteVergelijking) -> pd.DataFrame:
    """Per-categorie totals (IB + each offerte), then NRFP and grand totals
    — the same figures vergelijking_to_xlsx's Samenvatting sheet computes
    with formulas."""
    rows = []
    for cat in v.categorieen:
        regels = [r for r in v.regels if r.categorie == cat]
        ib = v.ib_categorie_totalen.get(cat, sum(r.ib_totaal for r in regels))
        offs = [sum(r.offerte_totaal(i) for r in regels) for i in range(len(v.namen))]
        rows.append((cat, ib, offs, True))
    excl_ib = sum(r[1] for r in rows)
    excl = [sum(r[2][i] for r in rows) for i in range(len(v.namen))]
    nrfp = [sum(it.totaal for it in items) for items in v.nrfp]
    rows.append(("NRFP / Buiten RFP", 0.0, nrfp, False))
    rows.append(("Totaal excl. NRFP", excl_ib, excl, True))
    rows.append(("Totaal incl. NRFP", excl_ib, [a + b for a, b in zip(excl, nrfp, strict=True)], True))

    out = []
    for cat, ib, offs, with_laagste in rows:
        rec = {"Categorie": cat, "IB (Budget)": ib}
        rec.update({n: t for n, t in zip(v.namen, offs, strict=True)})
        rec.update({f"{n} − IB": t - ib for n, t in zip(v.namen, offs, strict=True)})
        if len(v.namen) > 1:  # with one offerte it would read 'Alleen X' everywhere
            rec["Laagste"] = _laagste(v.namen, offs) if with_laagste else ""
        out.append(rec)
    return pd.DataFrame(out)


def unique_namen(namen: List[str]) -> List[str]:
    """Make offerte names unique for use as column headers — a second
    'Frimex' (e.g. a revised offerte) becomes 'Frimex (2)'."""
    out: List[str] = []
    for n in namen:
        naam, k = n, 2
        while naam in out:
            naam, k = f"{n} ({k})", k + 1
        out.append(naam)
    return out


def vergelijking_regels_df(v: OfferteVergelijking, totaal_rij: bool = True) -> pd.DataFrame:
    """Per-line table laid out like vergelijking_to_xlsx's Vergelijking sheet:
    IB aantal/prijs/totaal, then per offerte aantal/prijs/totaal/verschil
    t.o.v. IB, then the lowest offerte (only with 2+ offertes) and the
    match notes — plus a closing 'Totaal' row like the Excel's."""
    ranked = len(v.namen) > 1
    out = []
    for r in v.regels:
        rec = {"Categorie": r.categorie, "Art.Nr.Jumbo": r.art_nr_jumbo, "Omschrijving (offerte)": r.omschrijving,
               "IB omschrijving (Budget)": r.ib_omschrijving, "Eenheid": r.eenheid,
               "IB Aantal": r.ib_aantal, "IB Prijs p.e.": r.ib_prijs, "IB Totaal": r.ib_totaal}
        totalen = [r.offerte_totaal(i) for i in range(len(v.namen))]
        for i, n in enumerate(v.namen):
            aantal, prijs = r.offertes[i]
            rec[f"{n} Aantal"] = aantal
            rec[f"{n} Prijs p.e."] = prijs
            rec[f"{n} Totaal"] = totalen[i]
            rec[f"{n} − IB"] = totalen[i] - r.ib_totaal
        if ranked:
            rec["Laagste offerte"] = _laagste(v.namen, totalen)
        rec["Opmerking"] = r.opmerking
        out.append(rec)
    if totaal_rij and out:
        tot = {c: None for c in out[0]}
        tot["Categorie"] = "Totaal"
        for c in ["IB Totaal", *[f"{n} {k}" for n in v.namen for k in ("Totaal", "− IB")]]:
            tot[c] = sum(rec[c] for rec in out)
        if ranked:
            tot["Laagste offerte"] = _laagste(v.namen, [tot[f"{n} Totaal"] for n in v.namen])
        out.append(tot)
    return pd.DataFrame(out)


def _xl_laagste_formula(namen: List[str], refs: List[str]) -> str:
    """Excel formula naming the lowest offerte among those with a totaal > 0
    ('Alleen X' when only one quoted), mirroring _laagste."""
    m = "MIN(" + ",".join(f"IF({r}>0,{r},9E+99)" for r in refs) + ")"
    cnt = "+".join(f"({r}>0)" for r in refs)
    name = '""'
    for r, n in reversed(list(zip(refs, namen, strict=True))):
        name = f'IF({r}={m},"{n}",{name})'
    return f'=IF({cnt}=0,"",IF({cnt}=1,"Alleen "&{name},{name}))'


def vergelijking_to_xlsx(v: OfferteVergelijking, gemaakt_op: str = "") -> bytes:
    """Write the comparison as a formatted workbook (Samenvatting /
    Vergelijking / NRFP). Line and category totals are Excel formulas, so the
    sheet recalculates when someone edits an aantal or prijs; IB category
    totals taken from the Budget subheadings are values (blue, with a cell
    comment naming the Budget code)."""
    from openpyxl import Workbook
    from openpyxl.comments import Comment
    from openpyxl.formatting.rule import FormulaRule
    from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
    from openpyxl.utils import get_column_letter as L

    namen = [n.replace('"', "'") for n in v.namen]
    n_off = len(namen)
    EUR = '€ #,##0.00;-€ #,##0.00;"-"'
    NUM = '#,##0.##;-#,##0.##;"-"'
    PCT = '0.0%;-0.0%;"-"'
    f_norm = Font(name="Arial", size=10)
    f_bold = Font(name="Arial", size=10, bold=True)
    f_head = Font(name="Arial", size=10, bold=True, color="FFFFFF")
    f_title = Font(name="Arial", size=14, bold=True)
    f_input = Font(name="Arial", size=10, color="0000FF")
    f_note = Font(name="Arial", size=9, italic=True, color="666666")
    fill_head = PatternFill("solid", fgColor="1F4E78")
    fills_grp = [PatternFill("solid", fgColor=c) for c in ("D9D9D9", "DDEBF7", "FCE4D6", "E2EFDA", "FFF2CC", "EDE1F5")]
    fill_tot = PatternFill("solid", fgColor="D9E1F2")
    # Conditional-format fills need the colour as bgColor, hence start/end.
    fill_best = PatternFill(start_color="C6EFCE", end_color="C6EFCE", fill_type="solid")
    thin = Side(style="thin", color="BFBFBF")
    border = Border(top=thin, bottom=thin, left=thin, right=thin)
    wrap = Alignment(wrap_text=True, vertical="top")
    center = Alignment(horizontal="center", vertical="center", wrap_text=True)

    def head(cell, fill=fill_head, font=f_head):
        cell.font, cell.fill, cell.alignment, cell.border = font, fill, center, border

    def lowest_rule(col: str, row: int, cols: List[str]) -> FormulaRule:
        others = ",".join(f"IF(${o}{row}>0,${o}{row},9E+99)" for o in cols)
        return FormulaRule(formula=[f"AND({col}{row}>0,{col}{row}=MIN({others}))"], fill=fill_best)

    wb = Workbook()
    s = wb.active
    s.title = "Samenvatting"
    ws = wb.create_sheet("Vergelijking")
    wn = wb.create_sheet("NRFP")

    # ── Vergelijking ──
    ib_cols = (6, 7, 8)
    off_cols = [tuple(9 + 4 * i + k for k in range(4)) for i in range(n_off)]
    # 'Laagste offerte' and the green highlight only mean something with 2+
    # offertes — a single offerte vs IB would read 'Alleen X' on every line.
    ranked = n_off > 1
    col_best = 9 + 4 * n_off if ranked else None
    col_note = 9 + 4 * n_off + (1 if ranked else 0)
    ws.cell(1, 1, "Vergelijking per regel — Koel Installatie").font = f_title
    ws.cell(2, 1, "Regels met een aantal in minstens één offerte of in het IB, uitgelijnd op Art.Nr.Jumbo. "
                  "Totaal = Aantal × Prijs p.e. (materiaal + arbeid)"
                  + ("; groen = laagste offerte op die regel." if ranked else ".")).font = f_note
    for c, t in enumerate(["Categorie", "Art.Nr.Jumbo", "Omschrijving (offerte)", "IB omschrijving (Budget)",
                           "Eenheid"], 1):
        ws.merge_cells(start_row=4, start_column=c, end_row=5, end_column=c)
        head(ws.cell(4, c, t))
    groups = [("IB (Budget)", ib_cols, ("Aantal", "Prijs p.e.", "Totaal"))]
    groups += [(namen[i], off_cols[i], ("Aantal", "Prijs p.e.", "Totaal", "Verschil t.o.v. IB")) for i in range(n_off)]
    for gi, (label, cols, subs) in enumerate(groups):
        fill = fills_grp[gi % len(fills_grp)]
        ws.merge_cells(start_row=4, start_column=cols[0], end_row=4, end_column=cols[-1])
        ws.cell(4, cols[0], label)
        for c, t in zip(cols, subs, strict=True):
            head(ws.cell(4, c), fill, f_bold)
            head(ws.cell(5, c, t), fill, f_bold)
    for c, t in ([(col_best, "Laagste offerte")] if ranked else []) + [(col_note, "Opmerking")]:
        ws.merge_cells(start_row=4, start_column=c, end_row=5, end_column=c)
        head(ws.cell(4, c, t))

    first = 6
    r = first
    ia, ip, it = (L(c) for c in ib_cols)
    for regel in v.regels:
        for c, val in enumerate([regel.categorie, regel.art_nr_jumbo, regel.omschrijving,
                                 regel.ib_omschrijving, regel.eenheid], 1):
            ws.cell(r, c, val)
        ws.cell(r, ib_cols[0], regel.ib_aantal if _is_num(regel.ib_aantal) else None)
        ws.cell(r, ib_cols[1], regel.ib_prijs if _is_num(regel.ib_prijs) else None)
        ws.cell(r, ib_cols[2], f'=IF(OR({ia}{r}="",{ip}{r}=""),0,{ia}{r}*{ip}{r})')
        refs = []
        for i, (aantal, prijs) in enumerate(regel.offertes):
            ca, cp, ct, cv = off_cols[i]
            ws.cell(r, ca, aantal)
            ws.cell(r, cp, prijs)
            ws.cell(r, ct, f'=IF(OR({L(ca)}{r}="",{L(cp)}{r}=""),0,{L(ca)}{r}*{L(cp)}{r})')
            ws.cell(r, cv, f"={L(ct)}{r}-{it}{r}")
            refs.append(f"{L(ct)}{r}")
        if ranked:
            ws.cell(r, col_best, _xl_laagste_formula(namen, refs))
        ws.cell(r, col_note, regel.opmerking or None)
        for c in range(1, col_note + 1):
            cell = ws.cell(r, c)
            cell.font, cell.border, cell.alignment = f_norm, border, wrap
        for c in (ib_cols[0], *[oc[0] for oc in off_cols]):
            ws.cell(r, c).number_format = NUM
        for c in (ib_cols[1], ib_cols[2], *[x for oc in off_cols for x in oc[1:]]):
            ws.cell(r, c).number_format = EUR
        r += 1
    last = max(r - 1, first)
    ws.cell(r, 1, "Totaal")
    for c in (ib_cols[2], *[x for oc in off_cols for x in oc[2:]]):
        ws.cell(r, c, f"=SUM({L(c)}{first}:{L(c)}{last})").number_format = EUR
    for c in range(1, col_note + 1):
        ws.cell(r, c).font, ws.cell(r, c).fill, ws.cell(r, c).border = f_bold, fill_tot, border

    tot_cols = [L(oc[2]) for oc in off_cols]
    for col in tot_cols if ranked else []:
        ws.conditional_formatting.add(f"{col}{first}:{col}{last}", lowest_rule(col, first, tot_cols))
    widths = {1: 30, 2: 13, 3: 44, 4: 44, 5: 8, col_best: 16, col_note: 34}
    for c in range(1, col_note + 1):
        ws.column_dimensions[L(c)].width = widths.get(c, 13)
    ws.freeze_panes = ws.cell(first, 4)
    ws.auto_filter.ref = f"A5:{L(col_note)}{last}"

    # ── NRFP ──
    wn.cell(1, 1, "NRFP / Buiten RFP — posities niet in de RFP").font = f_title
    for c, t in enumerate(["Offerte", "Code", "Omschrijving", "Prijs", "Aantal", "Totaal"], 1):
        head(wn.cell(3, c, t))
    nr = 4
    for naam, items in zip(namen, v.nrfp, strict=True):
        for item in items:
            for c, val in enumerate([naam, item.code, item.omschrijving, item.prijs, item.aantal, item.totaal], 1):
                cell = wn.cell(nr, c, val)
                cell.font, cell.border = f_norm, border
            for c, fmt in ((4, EUR), (5, NUM), (6, EUR)):
                wn.cell(nr, c).number_format = fmt
            nr += 1
    nrfp_last = max(nr - 1, 4)
    for c, w in zip(range(1, 7), (14, 14, 60, 13, 9, 13), strict=True):
        wn.column_dimensions[L(c)].width = w
    wn.freeze_panes = "A4"

    # ── Samenvatting ──
    s.cell(1, 1, "Offertevergelijking Koel Installatie").font = f_title
    info = [("IB-werkboek", v.ib_bestand)] + [(f"Offerte {n}", b) for n, b in zip(namen, v.bestanden, strict=True)]
    if gemaakt_op:
        info.append(("Gemaakt op", gemaakt_op))
    for i, (k, val) in enumerate(info):
        s.cell(3 + i, 1, k).font = f_bold
        s.cell(3 + i, 2, val).font = f_norm
    hr = 4 + len(info)
    heads = (["Categorie", "IB (Budget)"] + namen + [f"{n} − IB" for n in namen]
             + [f"{n} t.o.v. IB" for n in namen] + (["Laagste"] if ranked else []))
    for c, t in enumerate(heads, 1):
        head(s.cell(hr, c, t))
    cat_rng = f"Vergelijking!$A${first}:$A${last}"
    sr = first_sum = hr + 1
    budget_rows = set()
    for cat in v.categorieen:
        s.cell(sr, 1, cat)
        if cat in v.ib_categorie_totalen:
            s.cell(sr, 2, v.ib_categorie_totalen[cat])
            s.cell(sr, 2).comment = Comment(
                f"Uit het IB Budget-tabblad, subtotaal-regel(s) {v.ib_categorie_codes.get(cat, '')} — "
                "zelfde bron als 'Totaal per categorie' in de app.", "Offerte Vergelijker")
            budget_rows.add(sr)
        else:
            s.cell(sr, 2, f"=SUMIFS(Vergelijking!${it}${first}:${it}${last},{cat_rng},$A{sr})")
        for i, col in enumerate(tot_cols):
            s.cell(sr, 3 + i, f"=SUMIFS(Vergelijking!${col}${first}:${col}${last},{cat_rng},$A{sr})")
        sr += 1
    last_cat = sr - 1
    nrfp_row = sr
    s.cell(sr, 1, "NRFP / Buiten RFP")
    s.cell(sr, 2, 0)
    s.cell(sr, 2).comment = Comment("Posities buiten de RFP staan niet in het IB.", "Offerte Vergelijker")
    for i, naam in enumerate(namen):
        s.cell(sr, 3 + i, f'=SUMIFS(NRFP!$F$4:$F${nrfp_last},NRFP!$A$4:$A${nrfp_last},"{naam}")')
    tot1, tot2 = sr + 1, sr + 2
    s.cell(tot1, 1, "Totaal excl. NRFP")
    s.cell(tot2, 1, "Totaal incl. NRFP")
    for c in range(2, 3 + n_off):
        s.cell(tot1, c, f"=SUM({L(c)}{first_sum}:{L(c)}{last_cat})")
        s.cell(tot2, c, f"={L(c)}{tot1}+{L(c)}{nrfp_row}")

    off_letters = [L(3 + i) for i in range(n_off)]
    for row in range(first_sum, tot2 + 1):
        for i, oc in enumerate(off_letters):
            s.cell(row, 3 + n_off + i, f"={oc}{row}-$B{row}")
            s.cell(row, 3 + 2 * n_off + i, f'=IF($B{row}=0,"",{oc}{row}/$B{row}-1)')
        if ranked and row != nrfp_row:  # fewer extras isn't a better price
            s.cell(row, 3 + 3 * n_off, _xl_laagste_formula(namen, [f"{oc}{row}" for oc in off_letters]))
            for oc in off_letters:
                s.conditional_formatting.add(f"{oc}{row}", lowest_rule(oc, row, off_letters))
        for c in range(1, len(heads) + 1):
            cell = s.cell(row, c)
            cell.font = f_input if (c == 2 and row in budget_rows) else (f_bold if row >= tot1 else f_norm)
            cell.border = border
            if 2 <= c < 3 + 2 * n_off:
                cell.number_format = EUR
            elif 3 + 2 * n_off <= c < 3 + 3 * n_off:
                cell.number_format = PCT
            if row >= tot1:
                cell.fill = fill_tot

    notes = [
        "Toelichting",
        "• IB (Budget): waar bekend het subtotaal van de bijbehorende Budget-regel (blauw, zie celopmerking); "
        "anders de som van de IB-regels op het tabblad Vergelijking.",
        "• Offertebedragen = Aantal × (Prijs materiaal + Arbeidskosten) per regel, alleen regels met een ingevuld aantal.",
        *(["• Groen = laagste offerte. 'Alleen …' = de andere partij(en) hebben hier niets geoffreerd "
           "(scope-verschil, geen prijsvoordeel)."] if ranked else []),
        "• Een verschil hoeft geen fout te zijn: partijen kunnen een andere variant, eenheid of scope hebben gekozen — "
        "controleer de regels op het tabblad Vergelijking.",
    ]
    for i, t in enumerate(notes):
        s.cell(tot2 + 2 + i, 1, t).font = f_bold if i == 0 else f_note
    s.column_dimensions["A"].width = 46
    for c in range(2, len(heads) + 1):
        s.column_dimensions[L(c)].width = 16
    s.freeze_panes = s.cell(hr + 1, 2)

    for sheet in (s, ws, wn):
        sheet.sheet_view.showGridLines = False
    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()
