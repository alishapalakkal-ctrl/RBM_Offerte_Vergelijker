"""Shared offerte-vs-IB comparison model (no Streamlit).

One `OfferteVergelijking` holds the lines of one or more offertes next to
the IB, per IB article. Each section builds it from its own matching
(matching_koeling.build_offerte_vergelijking, matching_aannemer.
build_aannemer_vergelijking, matching_sloopwerk.build_sloopwerk_vergelijking),
and every 'Alle resultaten' tab renders it the same way via
common.show_alle_resultaten — the layout of the Excel export's Vergelijking
sheet.
"""

import re
from dataclasses import dataclass, field
from typing import Dict, List, Optional

import pandas as pd


@dataclass
class VergelijkingRegel:
    """One line of the side-by-side comparison: an IB article (or, without
    an IB match, an offerte omschrijving) with each offerte's aantal and
    prijs p.e. — `offertes[i]` is (aantal, prijs) for offerte i, (None, None)
    when that offerte doesn't have the line at all."""
    categorie: Optional[str]
    art_nr_jumbo: Optional[str]
    #: the offerte's omschrijving — None for an IB-only line (see met_ib_only_regels)
    omschrijving: Optional[str]
    ib_omschrijving: Optional[str]
    eenheid: Optional[str]
    ib_aantal: Optional[float]
    ib_prijs: Optional[float]
    offertes: List[tuple]
    opmerking: str = ""
    #: The Budget row's own Totaal when the section uses it directly
    #: (Aannemer/Sloopwerk); None -> IB aantal × IB prijs.
    ib_totaal_budget: Optional[float] = None

    @property
    def ib_totaal(self) -> float:
        if self.ib_totaal_budget is not None:
            return self.ib_totaal_budget
        return (self.ib_aantal or 0) * (self.ib_prijs or 0)

    def offerte_totaal(self, i: int) -> float:
        aantal, prijs = self.offertes[i]
        return (aantal or 0) * (prijs or 0)


@dataclass
class OfferteVergelijking:
    namen: List[str]
    bestanden: List[str]
    regels: List[VergelijkingRegel]
    #: per offerte its NRFP/extra positions (objects with code, omschrijving,
    #: prijs, aantal, totaal — e.g. matching_koeling.NrfpItem)
    nrfp: List[list]
    #: Categorie labels in offerte order
    categorieen: List[str]
    #: Categorie -> Budget subheading Totaal where known
    ib_categorie_totalen: Dict[str, float] = field(default_factory=dict)
    ib_categorie_codes: Dict[str, str] = field(default_factory=dict)
    ib_bestand: str = ""


def categorieen_in_volgorde(regels: List[VergelijkingRegel]) -> List[str]:
    out: List[str] = []
    for r in regels:
        if r.categorie not in out:
            out.append(r.categorie)
    return out


def met_ib_only_regels(
    regels: List[VergelijkingRegel], ib_only: List[dict], code_naar_categorie: Dict[str, str], n_offertes: int = 1,
) -> List[VergelijkingRegel]:
    """`regels` plus a line per IB-only Budget row (matching_koeling.
    budget_ib_only_rows) — IB items no offerte line covers, e.g. a faywand
    type IB chose where the offerte quoted another, so the IB Totaal of the
    per-line view adds up to the Samenvatting's Totaal IB. Such a line has
    no offerte side (omschrijving None) and is filed under the offerte
    categorie of its Budget subtotal (via `code_naar_categorie`), else the
    subtotal's own description (e.g. 'Reiskosten'). Ordered by categorie —
    the offerte's categories first, each with its offerte lines before its
    IB-only lines."""
    extra = []
    for b in ib_only:
        categorie = code_naar_categorie.get(b["subcode"]) or b["subdesc"] or b["subcode"]
        extra.append(VergelijkingRegel(
            categorie=categorie, art_nr_jumbo=None, omschrijving=None, ib_omschrijving=b["desc"], eenheid=None,
            ib_aantal=b["aantal"], ib_prijs=b["prijs"], offertes=[(None, None)] * n_offertes,
            opmerking="alleen in IB (niet geoffreerd)" if b["row"] is not None
            else "alleen in IB — subtotaal zonder losse regels",
            ib_totaal_budget=b["totaal"],
        ))
    volgorde = categorieen_in_volgorde(regels + extra)
    alles = regels + extra  # stable sort keeps offerte lines before IB-only ones per categorie
    return sorted(alles, key=lambda r: volgorde.index(r.categorie))


def leverancier_uit_rijen(rows, max_rij: int = 40) -> Optional[str]:
    """The leverancier from a bidsheet's 'Prijsopgave Leverancier' line near
    the top (e.g. 'Van Wijnen', 'Van de Laar Sloopwerken') — the label sits
    in col 0 or col 1 depending on the template, the name in the next filled
    text cell. `rows` yields (row number, cells); None when there's no such
    line."""
    for r, cells in rows:
        if r > max_rij:
            break
        cells = list(cells)
        for i, c in enumerate(cells):
            if re.sub(r"[^a-z]", "", str(c or "").lower()) == "prijsopgaveleverancier":
                for v in cells[i + 1:]:
                    if isinstance(v, str) and v.strip():
                        return v.strip()[:30]
    return None


def match_opmerking(m) -> str:
    """Opmerking for a single offerte line from its Budget match (an object
    with .source and .score, e.g. matching_aannemer.BudgetMatch) — fuzzy
    and missing matches are the ones worth a second look."""
    if m is None:
        return "geen IB-match"
    if m.source == "fuzzy":
        return f"fuzzy match {m.score:.0f}%"
    return ""


def laagste_offerte(namen: List[str], totalen: List[float]) -> str:
    """Name of the lowest offerte among those that quoted (totaal > 0);
    'Alleen X' when only one did, '' when none did."""
    quoted = [(t, n) for t, n in zip(totalen, namen, strict=True) if t > 0]
    if not quoted:
        return ""
    best = min(quoted, key=lambda x: x[0])[1]
    return f"Alleen {best}" if len(quoted) == 1 else best


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
    """Per-line table laid out like the Excel export's Vergelijking sheet:
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
            rec["Laagste offerte"] = laagste_offerte(v.namen, totalen)
        rec["Opmerking"] = r.opmerking
        out.append(rec)
    if totaal_rij and out:
        tot = {c: None for c in out[0]}
        tot["Categorie"] = "Totaal"
        for c in ["IB Totaal", *[f"{n} {k}" for n in v.namen for k in ("Totaal", "− IB")]]:
            tot[c] = sum(rec[c] for rec in out)
        if ranked:
            tot["Laagste offerte"] = laagste_offerte(v.namen, [tot[f"{n} Totaal"] for n in v.namen])
        out.append(tot)
    return pd.DataFrame(out)
