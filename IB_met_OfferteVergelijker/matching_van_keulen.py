"""Van Keulen section: pure parsing/matching logic (no Streamlit).

Split out of pages/1_🟡_Van_Keulen.py so it can be unit-tested without a
Streamlit script-run context — importing that page module executes
st.set_page_config()/main() at import time.
"""

import re
from dataclasses import dataclass, field
from typing import Dict, List, Optional

import pandas as pd
import pdfplumber

try:
    from rapidfuzz import fuzz, process as rfprocess
    HAS_RAPIDFUZZ = True
except ImportError:
    HAS_RAPIDFUZZ = False


# ─── Data classes ──────────────────────────────────────────────────────────────

@dataclass
class PdfItem:
    art_nr: str
    description: str
    quantity: float
    unit_price: float
    total: float
    section: str = ""
    manual_ref: str = ""
    page: int = 0


@dataclass
class NettoItem:
    art_nr_jumbo: str
    manual_nr: str
    art_nr_leverancier: str
    groep: str
    description: str
    netto_price: float
    unit: str = ""


@dataclass
class IBItem:
    nummer: str
    description: str
    manual: str
    unit: str
    quantity: float
    price: float
    pdf_art_nr: str
    row: int = 0


@dataclass
class MatchResult:
    pdf_item: Optional[PdfItem] = None
    netto_item: Optional[NettoItem] = None
    ib_items: List[IBItem] = field(default_factory=list)
    pdf_match_method: str = "unmatched"
    ib_match_method: str = "unmatched"
    confidence: float = 0.0

    @property
    def ib_qty(self) -> Optional[float]:
        return sum(i.quantity for i in self.ib_items) if self.ib_items else None

    @property
    def price_diff(self) -> Optional[float]:
        if self.pdf_item and self.netto_item and self.netto_item.netto_price:
            return round(self.pdf_item.unit_price - self.netto_item.netto_price, 4)
        return None


# ─── PDF Parser ────────────────────────────────────────────────────────────────

_ITEM_END = re.compile(
    r'\s+((?:\d{1,3}\.)*\d+)\s+((?:\d{1,3}\.)*\d+,\d{2})\s+((?:\d{1,3}\.)*\d+,\d{2})\s*$'
)
_ART_NR  = re.compile(r'^(\d{3,7})\s+(.*)')
_MANUAL  = re.compile(r'^Manual\s+(\S+)', re.IGNORECASE)
_POSTAL  = re.compile(r'^\d{4}\s+[A-Z]{2}\s+')
_SECTION = re.compile(r'^[A-Za-zÀ-ÿ][A-Za-zÀ-ÿ /\-]{1,35}$')
_SKIP_SW = (
    'Offertenummer','Projectnummer','Ontvangende','Art.nr.','Art. nr.',
    'Op alle','Van Keulen','Postbus','www.','Pagina ',
    'Leverdatum','Uw contact','Planning','Aanvangsdatum',
    'EUR ','FSC ','PEFC ','ISO','Met vriendelijke',
    'Naam in blokletters','Handtekening','Totaalbedrag','Dit document',
    'Opdrachtgever','Jumbo Supermarkten','Mevr.','Dhr.','Nederland',
    'Offertedatum','Vervaldatum','Verkoper','Uw referentie',
)


def _dutch(s: str) -> float:
    return float(s.replace('.', '').replace(',', '.'))


def _skip(line: str) -> bool:
    if not line or len(line) <= 1:
        return True
    if _POSTAL.match(line):
        return True
    return any(line.startswith(s) for s in _SKIP_SW)


def parse_pdf(path) -> List[PdfItem]:
    items: List[PdfItem] = []
    section = manual_ref = art_nr = ""
    lines: List[str] = []
    page_nr = 0

    def flush():
        nonlocal art_nr, lines, manual_ref
        if not art_nr:
            return
        text = " ".join(lines).strip()
        m = _ITEM_END.search(text)
        if m:
            desc = text[:m.start()].strip()
            qty, price, total = _dutch(m.group(1)), _dutch(m.group(2)), _dutch(m.group(3))
            if price > 0:
                items.append(PdfItem(
                    art_nr=art_nr, description=desc, quantity=qty,
                    unit_price=price, total=total,
                    section=section, manual_ref=manual_ref, page=page_nr,
                ))
        art_nr = ""
        lines.clear()
        # A Manual code only applies to the single item it directly precedes —
        # clear it here so later items without their own "Manual ..." line
        # don't inherit a stale code from an earlier one.
        manual_ref = ""

    with pdfplumber.open(path) as pdf:
        for page_num, page in enumerate(pdf.pages, 1):
            if page_num == 1:
                continue
            page_nr = page_num
            text = page.extract_text()
            if not text:
                continue
            for raw in text.split('\n'):
                line = raw.strip()
                if _skip(line):
                    continue
                m = _MANUAL.match(line)
                if m:
                    flush()
                    manual_ref = m.group(1)
                    continue
                if line.startswith('['):
                    continue
                m = _ART_NR.match(line)
                if m:
                    if art_nr and not _ITEM_END.search(" ".join(lines)):
                        lines.append(line)
                        if _ITEM_END.search(" ".join(lines)):
                            flush()
                    else:
                        flush()
                        art_nr = m.group(1)
                        lines = [m.group(2)]
                        if _ITEM_END.search(lines[0]):  # price on same line as art.nr.
                            flush()
                    continue
                if art_nr:
                    lines.append(line)
                    if _ITEM_END.search(" ".join(lines)):
                        flush()
                elif line and _SECTION.match(line):
                    section = line
        flush()
    return items


# ─── Excel Readers ─────────────────────────────────────────────────────────────

def read_netto(src) -> List[NettoItem]:
    df = pd.read_excel(src, sheet_name=0, header=None)
    items: List[NettoItem] = []
    for _, row in df.iloc[2:].iterrows():
        def cell(i):
            return str(row.iloc[i]).strip() if pd.notna(row.iloc[i]) else ""
        jumbo = cell(0)
        if not jumbo or jumbo in ('nan', 'None'):
            continue
        art_lev = cell(2)
        if art_lev in ('nan', 'None', 'Geen artikelno.', 'Niet leverbaar'):
            art_lev = ""
        price = 0.0
        for ci in (16, 5):
            try:
                price = float(row.iloc[ci])
                break
            except (ValueError, TypeError):
                pass
        items.append(NettoItem(
            art_nr_jumbo=jumbo, manual_nr=cell(1), art_nr_leverancier=art_lev,
            groep=cell(3), description=cell(4), netto_price=price, unit=cell(11),
        ))
    return items


def read_ib_items(src, row_start: int = 1440, row_end: int = 2763) -> List[IBItem]:
    df = pd.read_excel(src, sheet_name="Budget", header=None, engine="openpyxl")
    items: List[IBItem] = []
    for idx in range(row_start - 1, min(row_end, len(df))):
        row = df.iloc[idx]
        def cell(i):
            return str(row.iloc[i]).strip() if i < len(row) and pd.notna(row.iloc[i]) else ""
        if cell(5) != "2":
            continue
        desc = cell(7)
        if not desc or desc in ("nan", "NaT", "0"):
            continue
        try:    qty   = float(row.iloc[10])
        except: qty   = 0.0
        try:    price = float(row.iloc[11])
        except: price = 0.0
        pdf_nr = cell(52) if len(row) > 52 else ""
        if pdf_nr in ("nan", "NaT"):
            pdf_nr = ""
        items.append(IBItem(
            nummer=cell(6), description=desc, manual=cell(8), unit=cell(9),
            quantity=qty, price=price, pdf_art_nr=pdf_nr, row=idx + 1,
        ))
    return items


# PDF art.nr. -> hardcoded IB Budget row. Bypasses the normal art.nr./fuzzy
# matching for cases where the correct IB row is known but doesn't match
# through the usual lookup (e.g. it's outside the standard row filter).
IB_ROW_OVERRIDE: Dict[str, int] = {
    "334363": 2612,
}


def read_ib_row_overrides(src, mapping: Dict[str, int] = IB_ROW_OVERRIDE) -> Dict[str, "IBItem"]:
    """Read specific Budget-sheet rows directly (ignoring the normal row
    filters) for article numbers with a hardcoded IB row mapping."""
    if not mapping:
        return {}
    df = pd.read_excel(src, sheet_name="Budget", header=None, engine="openpyxl")
    result: Dict[str, IBItem] = {}
    for art_nr, row_nr in mapping.items():
        idx = row_nr - 1
        if idx >= len(df):
            continue
        row = df.iloc[idx]
        def cell(i):
            return str(row.iloc[i]).strip() if i < len(row) and pd.notna(row.iloc[i]) else ""
        try:    qty   = float(row.iloc[10])
        except: qty   = 0.0
        try:    price = float(row.iloc[11])
        except: price = 0.0
        result[art_nr] = IBItem(
            nummer=cell(6), description=cell(7), manual=cell(8), unit=cell(9),
            quantity=qty, price=price, pdf_art_nr=art_nr, row=row_nr,
        )
    return result


# ─── Matcher ───────────────────────────────────────────────────────────────────

def build_matches(
    pdf_items: List[PdfItem],
    netto_items: List[NettoItem],
    ib_items: List[IBItem],
    fuzzy_threshold: int = 70,
    ib_row_override: Dict[str, IBItem] = None,
) -> List[MatchResult]:
    # A NETTO art.nr. leverancier can list multiple article numbers joined by
    # "+" (e.g. "492927+492928") meaning either one refers to this row.
    by_artnr: Dict[str, NettoItem] = {}
    for n in netto_items:
        if not n.art_nr_leverancier:
            continue
        for part in n.art_nr_leverancier.split('+'):
            part = part.strip()
            if part:
                by_artnr[part] = n
    by_manual: Dict[str, List[NettoItem]]  = {}
    for n in netto_items:
        k = n.manual_nr.strip().upper()
        if k:
            by_manual.setdefault(k, []).append(n)

    ib_by_pdf = {ii.pdf_art_nr: ii for ii in ib_items if ii.pdf_art_nr}
    ib_descs  = [ii.description for ii in ib_items]  if HAS_RAPIDFUZZ else []
    n_descs   = [n.description  for n  in netto_items] if HAS_RAPIDFUZZ else []

    results: List[MatchResult] = []
    for pi in pdf_items:
        mr = MatchResult(pdf_item=pi)

        ni = by_artnr.get(pi.art_nr)
        if ni:
            mr.netto_item = ni
            mr.pdf_match_method = "exact_artnr"
            mr.confidence = 1.0

        if not mr.netto_item and pi.manual_ref:
            cands = by_manual.get(pi.manual_ref.strip().upper(), [])
            if cands:
                if HAS_RAPIDFUZZ and len(cands) > 1:
                    # Multiple NETTO rows share this manual code — only accept the
                    # best description match if it clears the fuzzy threshold, so a
                    # weak/ambiguous pick doesn't silently attach the wrong row.
                    ds   = [c.description for c in cands]
                    best = rfprocess.extractOne(
                        pi.description, ds,
                        scorer=fuzz.token_set_ratio, score_cutoff=fuzzy_threshold,
                    )
                    if best:
                        mr.netto_item = cands[ds.index(best[0])]
                        mr.pdf_match_method = "manual_code"
                        mr.confidence = best[1] / 100.0
                else:
                    mr.netto_item = cands[0]
                    mr.pdf_match_method = "manual_code"
                    mr.confidence = 0.85

        if not mr.netto_item and HAS_RAPIDFUZZ:
            best = rfprocess.extractOne(
                pi.description, n_descs,
                scorer=fuzz.token_set_ratio, score_cutoff=fuzzy_threshold,
            )
            if best:
                mr.netto_item = netto_items[n_descs.index(best[0])]
                mr.pdf_match_method = "fuzzy"
                mr.confidence = best[1] / 100.0

        override = ib_row_override.get(pi.art_nr) if ib_row_override else None
        if override:
            mr.ib_items = [override]
            mr.ib_match_method = "manual_row"
            results.append(mr)
            continue

        ii = ib_by_pdf.get(pi.art_nr)
        if ii:
            mr.ib_items = [ii]
            mr.ib_match_method = "exact_artnr"
        elif HAS_RAPIDFUZZ and ib_descs:
            best = rfprocess.extractOne(
                pi.description, ib_descs,
                scorer=fuzz.token_set_ratio, score_cutoff=fuzzy_threshold,
            )
            if best:
                mr.ib_items = [ib_items[ib_descs.index(best[0])]]
                mr.ib_match_method = "fuzzy"

        results.append(mr)
    return results


# ─── Results → DataFrame ───────────────────────────────────────────────────────

def results_to_df(results: List[MatchResult]) -> pd.DataFrame:
    seen:  Dict[str, List[MatchResult]] = {}
    order: List[str] = []
    for mr in results:
        if not mr.pdf_item:
            continue
        k = mr.pdf_item.art_nr
        if k not in seen:
            seen[k] = []
            order.append(k)
        seen[k].append(mr)

    rows = []
    for k in order:
        mrs   = seen[k]
        first = mrs[0]
        pi    = first.pdf_item
        ni    = first.netto_item

        total_qty = sum(mr.pdf_item.quantity for mr in mrs)
        secs: List[str] = []
        for mr in mrs:
            s = mr.pdf_item.section if mr.pdf_item else ""
            if s and s not in secs:
                secs.append(s)

        ib_qty   = first.ib_qty
        qty_diff = round(total_qty - ib_qty, 4) if ib_qty is not None else None

        # Price diff only meaningful when art.nr. matched exactly in both lists
        if first.pdf_match_method == "exact_artnr" and first.price_diff is not None:
            price_diff = first.price_diff
            price_arrow = (
                f"↑ € {price_diff:,.2f}" if price_diff > 0.01
                else f"↓ € {abs(price_diff):,.2f}" if price_diff < -0.01
                else ""
            )
        else:
            price_diff  = None
            price_arrow = ""

        has_qty = qty_diff is not None and qty_diff != 0
        if   first.pdf_match_method == "unmatched": status = "unmatched"
        elif has_qty:                               status = "qty_diff"
        else:                                       status = "ok"

        rows.append({
            "_status":          status,
            "Methode":          first.pdf_match_method,   # kept for filters/metrics, hidden in table
            "Art.nr.":          pi.art_nr,
            "Manual nr.":       pi.manual_ref       if pi.manual_ref else "",
            "Omschrijving":     pi.description,
            "Sectie":           " / ".join(secs),
            "PDF Aantal":       total_qty,
            "PDF Prijs p.e.":   pi.unit_price,
            "PDF Totaal":       round(sum(mr.pdf_item.total for mr in mrs), 2),
            "IB Aantal":        ib_qty              if ib_qty  is not None else None,
            "Aantal verschil":  qty_diff            if qty_diff is not None else None,
            "NETTO Prijs p.e.": ni.netto_price      if ni else None,
            "Prijs verschil":   price_arrow,
            "Match %":          f"{first.confidence:.0%}" if first.confidence else "",
            "Pagina":           pi.page,
        })
    return pd.DataFrame(rows)


# ─── Budget Summary Rows ──────────────────────────────────────────────────────

# These IB rows are Lamellen Plafond items that actually belong to a different
# supplier — when present (aantal > 1) they must be flagged and left out of the
# Van Keulen IB total rather than silently mixed in.
LAMELLEN_PLAFOND_ROWS = {2263, 2264}


def read_budget_summary_rows(src, row_start: int, row_end: int) -> tuple[List[dict], List[dict]]:
    """Read rows row_start..row_end (1-based) from Budget sheet, keep rows where aantal >= 1.

    Returns (rows, excluded) where `excluded` holds the Lamellen Plafond rows
    that were skipped because they belong to a different supplier.
    """
    df = pd.read_excel(src, sheet_name="Budget", header=None, engine="openpyxl")
    rows = []
    excluded = []
    for rn in range(row_start, row_end + 1):
        idx = rn - 1
        if idx >= len(df):
            break
        row = df.iloc[idx]
        def cell(i):
            return str(row.iloc[i]).strip() if i < len(row) and pd.notna(row.iloc[i]) else ""
        try:    aantal = float(row.iloc[10])
        except: aantal = 0.0
        if aantal < 1:
            continue
        try:    prijs = float(row.iloc[11])
        except: continue
        if not pd.notna(prijs) or prijs <= 0:
            continue
        item = {
            "rij":         rn,
            "nummer":      cell(6),
            "artikelnaam": cell(7),
            "eenheid":     cell(9),
            "aantal":      aantal,
            "prijs":       prijs,
            "totaal":      round(aantal * prijs, 2),
        }
        if rn in LAMELLEN_PLAFOND_ROWS and aantal > 1:
            excluded.append(item)
            continue
        rows.append(item)
    return rows, excluded
