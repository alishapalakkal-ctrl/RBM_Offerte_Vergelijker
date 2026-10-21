"""Sloopwerk section: pure parsing/matching logic (no Streamlit).

Compares a demolition contractor's offerte (e.g. Van de Laar 'vertaling
bidsheet' — an Excel workbook, not a PDF like Aannemer/Koeling) against the
IB workbook's Budget sheet, section A.01 (Sloopwerk). Like Aannemer, the
Budget rows here have no numeric-prefix bridge to join against — matching is
purely description-based (fuzzy), restricted per Groep/categorie via
data/sloopwerk_categorie_budget_codes.csv so unrelated categories can't
cross-match (verified against the real Poeldijk files: matching without that
restriction produces nonsense hits for generic-sounding descriptions).

Split out of pages/4_🚧_Sloopwerk.py so it can be unit-tested without a
Streamlit script-run context — importing that page module executes
st.set_page_config()/session_state code at import time.
"""

import io
from dataclasses import dataclass
from pathlib import Path
from typing import Dict, List, Optional

import pandas as pd
import pdfplumber

from matching_koeling import (
    HAS_RAPIDFUZZ,
    _budget_code_totals,
    _budget_prijs_totaal,
    _budget_row,
    _nh,
    _norm_text,
    _pdf_cell,
    build_categorie_totalen,
    load_budget_df,
)
from vergelijking import (
    OfferteVergelijking,
    VergelijkingRegel,
    categorieen_in_volgorde,
    leverancier_uit_rijen,
    match_opmerking,
)

if HAS_RAPIDFUZZ:
    from rapidfuzz import fuzz

__all__ = [
    "HAS_RAPIDFUZZ",
    "BUDGET_SLOOPWERK_VANDELAAR_ROW_END",
    "BUDGET_SLOOPWERK_VANDELAAR_ROW_START",
    "OFFERTE_SHEET_NAME",
    "BudgetMatch",
    "OfferteItem",
    "build_categorie_totalen",
    "build_sloopwerk_df",
    "build_sloopwerk_vergelijking",
    "get_sloopwerk_totaal_ib",
    "load_budget_df",
    "load_categorie_budget_codes",
    "match_sloopwerk_budget",
    "parse_pdf",
    "read_offerte",
    "read_offerte_file",
]


@dataclass
class OfferteItem:
    row: int
    groep: str  # "Algemeen", "Vergunningen", ... or "NRFP"
    art_nr_jumbo: Optional[str]
    manual_nr: Optional[str]
    art_nr_leverancier: Optional[str]
    omschrijving: str
    eenheid: Optional[str]
    aantal: Optional[float]
    totaal: Optional[float]  # offerte's own "Waarde offerte Totaal"


@dataclass
class BudgetMatch:
    budget_row: int
    code: Optional[str]
    omschrijving: str
    aantal: Optional[float]
    prijs: Optional[float]
    totaal: Optional[float]
    score: float
    source: str  # "fuzzy"


#: Budget-sheet rows covering the Sloopwerk (Van de Laar) 'A.01' section —
#: row 18 is the first subheading (A.01.01 Algemeen), row 272 is the last
#: leaf row before 'A.02 Aannemerswerk' starts at row 273 (verified against
#: the real Poeldijk Demolition workbook).
BUDGET_SLOOPWERK_VANDELAAR_ROW_START = 18
BUDGET_SLOOPWERK_VANDELAAR_ROW_END = 272

#: The offerte workbook's single data sheet (verified against the real
#: P260208-B Poeldijk 'Offerte Demontage' file) — unlike Koeling, there's
#: only ever one sheet name for this leverancier's format, so it's a
#: constant rather than a per-call parameter.
OFFERTE_SHEET_NAME = "vertaling bidsheet - offerte"

#: 1-based columns in the offerte's item table (0-based index in parentheses
#: matches pandas' header=None column position) — the header row repeats
#: every 5-8 rows above each Groep block, so items are found by scanning
#: for a filled Groep + Omschrijving pair rather than a fixed start row.
_COL_ART_NR_JUMBO = 1
_COL_MANUAL_NR = 2
_COL_ART_NR_LEVERANCIER = 3
_COL_GROEP = 4
_COL_OMSCHRIJVING = 5
_COL_EENHEID = 7
_COL_AANTAL = 9
_COL_TOTAAL = 21


def read_offerte(file_bytes: bytes, sheet_name: str = OFFERTE_SHEET_NAME) -> List[OfferteItem]:
    try:
        df = pd.read_excel(io.BytesIO(file_bytes), sheet_name=sheet_name, header=None, engine="openpyxl")
    except ValueError:
        return []

    def cell(row, col):
        return row.iloc[col] if len(row) > col else None

    items: List[OfferteItem] = []
    for idx, row in df.iterrows():
        omschrijving = cell(row, _COL_OMSCHRIJVING)
        groep = cell(row, _COL_GROEP)
        if not isinstance(omschrijving, str) or not omschrijving.strip():
            continue
        if omschrijving.strip() == "Omschrijving Item":  # repeated header row
            continue
        if not isinstance(groep, str) or not groep.strip():
            continue

        art_nr_jumbo = cell(row, _COL_ART_NR_JUMBO)
        manual_nr = cell(row, _COL_MANUAL_NR)
        art_nr_lev = cell(row, _COL_ART_NR_LEVERANCIER)
        eenheid = cell(row, _COL_EENHEID)
        aantal = cell(row, _COL_AANTAL)
        totaal = cell(row, _COL_TOTAAL)

        items.append(OfferteItem(
            row=int(idx) + 1,
            groep=groep.strip(),
            art_nr_jumbo=str(art_nr_jumbo).strip() if pd.notna(art_nr_jumbo) else None,
            manual_nr=str(manual_nr).strip() if pd.notna(manual_nr) else None,
            art_nr_leverancier=str(art_nr_lev).strip() if pd.notna(art_nr_lev) else None,
            omschrijving=omschrijving.strip(),
            eenheid=str(eenheid).strip() if isinstance(eenheid, str) else None,
            aantal=float(aantal) if pd.notna(aantal) else None,
            totaal=float(totaal) if pd.notna(totaal) else None,
        ))
    return items


#: Header labels (normalized: lowercase, alnum only) of the item table in a
#: PDF print of the offerte. The print hides columns the Excel has (Manual
#: nr, Art. nr. leverancier), so unlike read_offerte's fixed _COL_*
#: positions, the PDF's columns are found by name per header row.
_PDF_HEADER_KEYS = {
    "art_nr_jumbo": lambda n: n == "artnrjumbo",
    "groep": lambda n: n == "groep",
    "omschrijving": lambda n: n == "omschrijvingitem",
    "eenheid": lambda n: n == "eenheid",
    "aantal": lambda n: n == "aantalpereenheid",
    "totaal": lambda n: n.startswith("waardeoffertetot"),
}


def _parse_pdf_rows(rows) -> List[OfferteItem]:
    """Item rows of a PDF-printed offerte (e.g. Van de Laar's 'Demontage
    template offerte en eindafrekening', Rotterdam P250324): each table
    header row sets the column mapping, then a row counts when Groep and
    Omschrijving are both filled — the same rule as read_offerte. Tables
    without a Groep column (the eindafrekening half, summary boxes) never
    set a mapping; their rows are also narrower than the item table, which
    is how they're kept out when they follow it — a continuation of the
    item table on the next page has the same width and is still read."""
    items: List[OfferteItem] = []
    cols: Dict[str, int] = {}
    width = 0
    for r, raw in rows:
        normed = [_nh(v) for v in raw]
        if normed and normed[0] == "artnrjumbo":
            cols = {}
            for key, test in _PDF_HEADER_KEYS.items():
                for c, n in enumerate(normed):
                    if test(n):
                        cols[key] = c
                        break
            if "groep" not in cols or "omschrijving" not in cols:
                cols = {}
            width = len(raw)
            continue
        if not cols or len(raw) != width:
            continue

        def cell(key):
            c = cols.get(key)
            return _pdf_cell(raw[c]) if c is not None and c < len(raw) else None

        groep, omschrijving = cell("groep"), cell("omschrijving")
        if not isinstance(groep, str) or not isinstance(omschrijving, str):
            continue
        art_nr, aantal, totaal, eenheid = cell("art_nr_jumbo"), cell("aantal"), cell("totaal"), cell("eenheid")
        items.append(OfferteItem(
            row=r, groep=groep, art_nr_jumbo=str(art_nr) if art_nr is not None else None,
            manual_nr=None, art_nr_leverancier=None, omschrijving=omschrijving,
            eenheid=eenheid if isinstance(eenheid, str) else None,
            aantal=aantal if isinstance(aantal, float) else None,
            totaal=totaal if isinstance(totaal, float) else None,
        ))
    return items


def _pdf_table_rows(file_bytes: bytes):
    r = 0
    with pdfplumber.open(io.BytesIO(file_bytes)) as pdf:
        for page in pdf.pages:
            for table in page.extract_tables():
                for raw in table:
                    r += 1
                    yield r, raw


def parse_pdf(file_bytes: bytes) -> List[OfferteItem]:
    """Same as read_offerte, for the offerte printed to PDF."""
    return _parse_pdf_rows(_pdf_table_rows(file_bytes))


def read_offerte_file(file_bytes: bytes, filename: str) -> tuple:
    """(offerte items, leverancier) from the offerte as Excel or as PDF —
    the leverancier from its 'Prijsopgave Leverancier' line, else 'Offerte'."""
    if filename.lower().endswith(".pdf"):
        return parse_pdf(file_bytes), leverancier_uit_rijen(_pdf_table_rows(file_bytes)) or "Offerte"
    try:
        df = pd.read_excel(io.BytesIO(file_bytes), sheet_name=OFFERTE_SHEET_NAME, header=None,
                           engine="openpyxl", nrows=40)
        rows = enumerate(df.values.tolist(), start=1)
        leverancier = leverancier_uit_rijen(rows)
    except ValueError:
        leverancier = None
    return read_offerte(file_bytes), leverancier or "Offerte"


ART_NR_MAPPING_DIR = Path(__file__).parent / "data"


def load_categorie_budget_codes() -> Dict[str, List[tuple]]:
    """Read data/sloopwerk_categorie_budget_codes.csv: offerte 'Groep' label
    -> [(budget_code, budget_omschrijving), ...]. Keyed by the offerte's own
    raw Groep string so it matches OfferteItem.groep verbatim — same
    convention as Aannemer/Koeling's load_categorie_budget_codes, whose
    result feeds directly into build_categorie_totalen (reused as-is)."""
    path = ART_NR_MAPPING_DIR / "sloopwerk_categorie_budget_codes.csv"
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


def get_sloopwerk_totaal_ib(
    bdf: Optional[pd.DataFrame],
    row_start: int = BUDGET_SLOOPWERK_VANDELAAR_ROW_START, row_end: int = BUDGET_SLOOPWERK_VANDELAAR_ROW_END,
) -> Optional[float]:
    """Sum every A.01.0X subheading's own Totaal directly (Algemeen,
    Reiskosten, ... Hergebruik/recycling — all 16) — the whole-section IB
    figure for the Samenvatting's top 'Totaal IB' KPI, independent of
    build_categorie_totalen's per-Groep rollup (which only covers categories
    the offerte itself quotes). Same approach as Aannemer's
    get_aannemer_totaal_ib — equals the 'A.01 Sloopwerk' section's own total,
    computed explicitly from its children rather than trusting one cell."""
    if bdf is None:
        return None
    code_totals = _budget_code_totals(bdf, row_start, row_end)
    total = 0.0
    found = False
    for code, (_desc, _aantal, totaal) in code_totals.items():
        if not isinstance(code, str) or not code.startswith("A.01.") or code.count(".") != 2:
            continue
        if totaal:
            total += totaal
            found = True
    return total if found else None


def match_sloopwerk_budget(
    offerte: List[OfferteItem], bdf: Optional[pd.DataFrame], categorie_mapping: Dict[str, List[tuple]],
    row_start: int = BUDGET_SLOOPWERK_VANDELAAR_ROW_START, row_end: int = BUDGET_SLOOPWERK_VANDELAAR_ROW_END,
    fuzzy_threshold: int = 40,
) -> Dict[int, BudgetMatch]:
    """Offerte-row -> Budget-row join via rapidfuzz description matching,
    restricted to Budget leaf rows in the same categorie as the offerte
    item's own Groep (via categorie_mapping). Items whose Groep has no entry
    in categorie_mapping (NRFP, or any category not yet mapped) are left
    unmatched rather than fuzzy-matched against the whole section — verified
    against the real Poeldijk files that doing so produces nonsense matches
    for generically-worded descriptions (e.g. "Toeslag uurloon avondwerk"
    matching "Demonteren en afvoeren keukenblok")."""
    if bdf is None:
        return {}

    code_to_categorie: Dict[str, str] = {}
    for categorie, entries in categorie_mapping.items():
        for code, _omschrijving in entries:
            code_to_categorie[code] = categorie

    candidates = []
    for r in range(row_start, row_end + 1):
        result = _budget_row(bdf, r)
        if result is None:
            continue
        desc, aantal, row = result
        code = row.iloc[6] if len(row) > 6 else None
        # Leaf item rows are 4-dot outline codes (e.g. 'A.01.01.02');
        # group-header/subtotal rows ('A.01.01', 'A.01') have no individual
        # item to join against.
        if not isinstance(code, str) or code.count(".") != 3:
            continue
        prijs, totaal = _budget_prijs_totaal(row)
        prefix = ".".join(code.split(".")[:3])
        candidates.append({
            "row": r, "desc": desc, "aantal": aantal, "prijs": prijs, "totaal": totaal,
            "code": code, "categorie": code_to_categorie.get(prefix),
        })

    used_budget = set()
    matches: Dict[int, BudgetMatch] = {}

    def _claim(i, c, score):
        used_budget.add(c["row"])
        matches[i] = BudgetMatch(
            budget_row=c["row"], code=c["code"], omschrijving=c["desc"], aantal=c["aantal"],
            prijs=c["prijs"], totaal=c["totaal"], score=score, source="fuzzy",
        )

    matchable = [
        (i, o) for i, o in enumerate(offerte)
        if o.groep != "NRFP" and o.groep in categorie_mapping and o.omschrijving
    ]

    if HAS_RAPIDFUZZ:
        pairs = []
        for i, o in matchable:
            pool = [c for c in candidates if c["categorie"] == o.groep]
            for c in pool:
                if c["row"] in used_budget:
                    continue
                score = fuzz.token_set_ratio(_norm_text(o.omschrijving), _norm_text(c["desc"]))
                if score >= fuzzy_threshold:
                    pairs.append((score, i, c))
        pairs.sort(key=lambda x: -x[0])
        for score, i, c in pairs:
            if i in matches or c["row"] in used_budget:
                continue
            _claim(i, c, float(score))

    return matches


def build_sloopwerk_df(offerte: List[OfferteItem], matches: Dict[int, BudgetMatch]) -> pd.DataFrame:
    rows = []
    for i, o in enumerate(offerte):
        m = matches.get(i)

        aantal_o = o.aantal or 0
        aantal_i = m.aantal if (m and m.aantal is not None) else None
        aantal_diff = (aantal_o - aantal_i) if aantal_i is not None else None

        offerte_totaal = o.totaal
        if m and m.totaal is not None:
            ib_totaal = m.totaal
        elif m and m.prijs is not None and aantal_o:
            ib_totaal = m.prijs * aantal_o
        else:
            ib_totaal = None

        totaal_diff = (offerte_totaal - ib_totaal) if (offerte_totaal is not None and ib_totaal is not None) else None
        totaal_diff_pct = (totaal_diff / ib_totaal * 100) if (totaal_diff is not None and ib_totaal) else None

        if aantal_diff is None or aantal_diff == 0:
            significant_aantal_diff = False
        elif aantal_i:
            significant_aantal_diff = abs(aantal_diff) / aantal_i * 100 > 10
        else:
            significant_aantal_diff = True

        if o.groep == "NRFP":
            status = "nrfp"
        elif m is None:
            status = "unmatched"
        elif totaal_diff is not None and abs(totaal_diff) > 0.01:
            status = "prijs_afwijking"
        elif significant_aantal_diff:
            status = "aantal_afwijking"
        else:
            status = "ok"

        rows.append({
            "_status": status,
            "_offerte_totaal": offerte_totaal,
            "_ib_totaal": ib_totaal,
            "Groep": o.groep,
            "Offerte omschrijving": o.omschrijving,
            "IB omschrijving": m.omschrijving if m else None,
            "Match": m.source if m else "geen match",
            "Match %": f"{m.score:.0f}%" if m else "",
            "Eenheid": o.eenheid,
            "Offerte Aantal": aantal_o if aantal_o else None,
            "IB Aantal": aantal_i,
            "Aantal verschil": aantal_diff,
            "Offerte Totaal": offerte_totaal,
            "IB Totaal": ib_totaal,
            "Totaal verschil (€)": totaal_diff,
            "Totaal verschil (%)": totaal_diff_pct,
        })
    return pd.DataFrame(rows)


def build_sloopwerk_vergelijking(
    offerte: List[OfferteItem], matches: Dict[int, BudgetMatch],
    categorie_totalen: Optional[Dict[str, float]] = None, naam: str = "Van de Laar", bestand: str = "",
) -> OfferteVergelijking:
    """The offerte next to the IB per line, for the shared 'Alle resultaten'
    layout (common.show_alle_resultaten) — lines with an aantal, NRFP left
    out. The offerte only has a line total ('Waarde offerte Totaal'), so its
    prijs p.e. is totaal / aantal; IB totaal is the Budget row's own Totaal
    (IB prijs × offerte aantal when it has none), as in build_sloopwerk_df."""
    regels = []
    for i, o in enumerate(offerte):
        if o.groep == "NRFP" or not o.aantal:
            continue
        m = matches.get(i)
        if m and m.totaal is not None:
            ib_totaal = m.totaal
        elif m and m.prijs is not None:
            ib_totaal = m.prijs * o.aantal
        else:
            ib_totaal = None
        regels.append(VergelijkingRegel(
            categorie=o.groep, art_nr_jumbo=o.art_nr_jumbo, omschrijving=o.omschrijving,
            ib_omschrijving=m.omschrijving if m else None, eenheid=o.eenheid,
            ib_aantal=m.aantal if m else None, ib_prijs=m.prijs if m else None, ib_totaal_budget=ib_totaal,
            offertes=[(o.aantal, o.totaal / o.aantal if o.totaal is not None else None)],
            opmerking=match_opmerking(m),
        ))
    return OfferteVergelijking(
        namen=[naam], bestanden=[bestand], regels=regels, nrfp=[[]],
        categorieen=categorieen_in_volgorde(regels), ib_categorie_totalen=dict(categorie_totalen or {}),
    )
