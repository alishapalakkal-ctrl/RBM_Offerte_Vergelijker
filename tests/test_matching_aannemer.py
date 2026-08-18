import pandas as pd
import pytest

from matching_aannemer import (
    HAS_RAPIDFUZZ,
    BudgetMatch,
    OfferteItem,
    _dutch_number,
    build_aannemer_df,
    match_aannemer_budget,
)

# NOTE: parse_pdf itself is not unit-tested here — pdfplumber's
# extract_tables() output can only meaningfully be validated against a real
# PDF (verified this session against the actual Van Wijnen bidsheet), not a
# hand-built stand-in. _dutch_number and the matching/df-building logic
# below are the parts that don't depend on pdfplumber at all.


# ─── numeric parsing ─────────────────────────────────────────────────────────

@pytest.mark.parametrize("raw,expected", [
    (None, None),
    ("", None),
    ("-", None),
    ("€ 1 07.484,15", 107484.15),   # internal-space quirk seen in the real PDF
    ("€ 9 1,36", 91.36),
    ("9,2", 9.2),
    ("3", 3.0),
])
def test_dutch_number(raw, expected):
    assert _dutch_number(raw) == expected


# ─── matching ────────────────────────────────────────────────────────────────

def _offerte_item(**overrides) -> OfferteItem:
    base = dict(
        row=1, pos="Pos 1 Kozijnen en deuren", art_nr_jumbo="jum-AA-0001", manual_nr=None,
        art_nr_leverancier=None, omschrijving="Houten binnendeurkozijn t.b.v. deur 930 x 2.315 mm",
        toelichting=None, eenheid="Stuk", aantal=1.0, materiaalkosten_pe=83.48, normtijd_min_pe=45.0,
        tarief_uur=61.36, prijs_pe=129.5, totale_uren=None, totale_materiaalkosten=None,
        arbeid=None, totaal=None, opmerkingen=None, page=1,
    )
    base.update(overrides)
    return OfferteItem(**base)


def _budget_df(rows: dict) -> pd.DataFrame:
    """rows: {1-based row number: (code, desc, aantal, prijs, totaal)} -> a
    13-column frame shaped like the real Budget sheet
    (code=G/6, desc=H/7, aantal=K/10, prijs=L/11, totaal=M/12)."""
    max_row = max(rows)
    data = [[None] * 13 for _ in range(max_row)]
    for r, (code, desc, aantal, prijs, totaal) in rows.items():
        data[r - 1][6] = code
        data[r - 1][7] = desc
        data[r - 1][10] = aantal
        data[r - 1][11] = prijs
        data[r - 1][12] = totaal
    return pd.DataFrame(data)


def test_match_aannemer_budget_none_bdf_returns_empty():
    assert match_aannemer_budget([_offerte_item()], None, {}) == {}


def test_match_aannemer_budget_excludes_group_header_rows():
    # Only 3-dot leaf codes are matchable; a 2-dot group-header row like
    # 'A.02.04' ("Kozijnen, ramen en deuren") must never be a candidate,
    # even though its description could otherwise fuzzy-match something.
    offerte = [_offerte_item(omschrijving="Kozijnen, ramen en deuren")]
    bdf = _budget_df({298: ("A.02.04", "Kozijnen, ramen en deuren", 0, None, 1145.76)})

    matches = match_aannemer_budget(offerte, bdf, {}, row_start=298, row_end=298)

    assert matches == {}


def test_match_aannemer_budget_skips_nrfp_rows():
    offerte = [_offerte_item(pos="NRFP", art_nr_jumbo="NRFP-0001", omschrijving="Precariokosten")]
    bdf = _budget_df({299: ("A.02.04.01", "Houten binnendeurkozijn t.b.v. deur 930 x 2.315 mm", 0, 100.0, 0.0)})

    matches = match_aannemer_budget(offerte, bdf, {}, row_start=299, row_end=299)

    assert matches == {}


def test_match_aannemer_budget_csv_override_takes_priority():
    offerte = [_offerte_item()]
    bdf = _budget_df({
        299: ("A.02.04.01", "Houten binnendeurkozijn t.b.v. deur 930 x 2.315 mm", 0, 134.61, 0.0),
    })
    mapping = {"jum-AA-0001": "Houten binnendeurkozijn t.b.v. deur 930 x 2.315 mm"}

    matches = match_aannemer_budget(offerte, bdf, mapping, row_start=299, row_end=299)

    assert matches[0].source == "handmatig"
    assert matches[0].score == 100.0
    assert matches[0].budget_row == 299


@pytest.mark.skipif(not HAS_RAPIDFUZZ, reason="rapidfuzz not installed")
def test_match_aannemer_budget_fuzzy_fallback():
    offerte = [_offerte_item(art_nr_jumbo="jum-AA-9999")]  # no CSV entry
    bdf = _budget_df({
        299: ("A.02.04.01", "Houten binnendeurkozijn t.b.v. deur 930 x 2.315 mm", 0, 134.61, 0.0),
    })

    matches = match_aannemer_budget(offerte, bdf, {}, row_start=299, row_end=299)

    assert matches[0].source == "fuzzy"
    assert matches[0].budget_row == 299


@pytest.mark.skipif(not HAS_RAPIDFUZZ, reason="rapidfuzz not installed")
def test_match_aannemer_budget_falls_back_when_matched_row_has_no_aantal():
    # CSV override lands on a row with no aantal — IB actually selected a
    # different but similarly-described gevelwand variant on another row.
    offerte = [_offerte_item(art_nr_jumbo="jum-AA-0001", omschrijving="Gevelwand type A")]
    bdf = _budget_df({
        299: ("A.02.04.01", "Gevelwand type A", 0, 100.0, 0.0),
        300: ("A.02.04.02", "Gevelwand type B", 4.0, 110.0, 440.0),
    })
    mapping = {"jum-AA-0001": "Gevelwand type A"}

    matches = match_aannemer_budget(offerte, bdf, mapping, row_start=299, row_end=300)

    assert matches[0].budget_row == 300
    assert matches[0].aantal == 4.0
    assert matches[0].source == "fuzzy"


def test_match_aannemer_budget_no_double_claim():
    offerte = [
        _offerte_item(row=1, art_nr_jumbo="jum-AA-0001", omschrijving="Houten binnendeurkozijn"),
        _offerte_item(row=2, art_nr_jumbo="jum-AA-0002", omschrijving="Houten binnendeurkozijn"),
    ]
    bdf = _budget_df({299: ("A.02.04.01", "Houten binnendeurkozijn", 0, 134.61, 0.0)})

    matches = match_aannemer_budget(offerte, bdf, {}, row_start=299, row_end=299)

    assert len(matches) == 1


# ─── df building / status ───────────────────────────────────────────────────

def test_build_aannemer_df_flags_price_and_quantity_differences():
    offerte = [
        _offerte_item(row=1, prijs_pe=100.0, aantal=2),   # price mismatch
        _offerte_item(row=2, prijs_pe=100.0, aantal=3),   # qty mismatch
        _offerte_item(row=3, prijs_pe=100.0, aantal=2),   # ok
        _offerte_item(row=4),                             # unmatched
        _offerte_item(row=5, pos="NRFP"),                 # nrfp
    ]
    match_row = BudgetMatch(budget_row=299, code="A.02.04.01", omschrijving="x",
                             aantal=2, prijs=90.0, totaal=180.0, score=100.0, source="handmatig")
    matches = {
        0: match_row,
        1: BudgetMatch(budget_row=300, code="A.02.04.02", omschrijving="x",
                       aantal=2, prijs=100.0, totaal=200.0, score=100.0, source="handmatig"),
        2: BudgetMatch(budget_row=301, code="A.02.04.03", omschrijving="x",
                       aantal=2, prijs=100.0, totaal=200.0, score=100.0, source="handmatig"),
    }

    df = build_aannemer_df(offerte, matches)

    assert list(df["_status"]) == ["prijs_afwijking", "aantal_afwijking", "ok", "unmatched", "nrfp"]


def test_build_aannemer_df_none_match_is_unmatched():
    offerte = [_offerte_item()]

    df = build_aannemer_df(offerte, {})

    assert df.loc[0, "_status"] == "unmatched"
    assert df.loc[0, "IB omschrijving"] is None
