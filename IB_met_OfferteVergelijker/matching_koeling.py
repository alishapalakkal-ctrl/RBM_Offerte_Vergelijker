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
    "aantal": {"aantal"},
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


def parse_offerte(file_bytes: bytes, sheet_name: str) -> List[OfferteItem]:
    items: List[OfferteItem] = []
    category = None
    header_map: Dict[str, int] = {}

    for r, row_vals in _rows_from_bytes(file_bytes, sheet_name):
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

        items.append(OfferteItem(
            row=r, categorie=category, subgroep=c1 if isinstance(c1, str) else None,
            omschrijving=desc.strip(), eenheid=get("eenheid"),
            prijs_materiaal=get("prijs_materiaal"), arbeid=get("arbeid"),
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
        df = pd.read_excel(io.BytesIO(file_bytes), sheet_name=sheet_name, header=None, engine="openpyxl")
    except ValueError:
        return []

    items: List[NrfpItem] = []
    for _, row in df.iterrows():
        code = row.iloc[0] if len(row) > 0 else None
        if not isinstance(code, str) or not _NRFP_CODE_RE.match(code.strip()):
            continue
        totaal = row.iloc[4] if len(row) > 4 else None
        if not (isinstance(totaal, (int, float)) and pd.notna(totaal) and totaal):
            continue
        omschrijving = row.iloc[1] if len(row) > 1 else None
        prijs = row.iloc[2] if len(row) > 2 else None
        aantal = row.iloc[3] if len(row) > 3 else None
        items.append(NrfpItem(
            code=code.strip(),
            omschrijving=str(omschrijving).strip() if isinstance(omschrijving, str) else None,
            prijs=prijs if pd.notna(prijs) else None,
            aantal=aantal if pd.notna(aantal) else None,
            totaal=float(totaal),
        ))
    return items


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
