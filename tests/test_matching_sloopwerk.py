import io

import openpyxl
import pandas as pd
import pytest

from matching_sloopwerk import (
    HAS_RAPIDFUZZ,
    BudgetMatch,
    OfferteItem,
    build_sloopwerk_df,
    build_sloopwerk_vergelijking,
    get_sloopwerk_totaal_ib,
    match_sloopwerk_budget,
    read_offerte,
)


def _xlsx_bytes(sheet_name: str, rows: list) -> bytes:
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = sheet_name
    for row in rows:
        ws.append(row)
    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()


def _offerte_item(**overrides) -> OfferteItem:
    base = dict(
        row=1, groep="Algemeen", art_nr_jumbo="No. 1", manual_nr=None, art_nr_leverancier=None,
        omschrijving="KLIC-melding", eenheid="Post", aantal=1.0, totaal=69.34,
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


# ─── offerte reading ──────────────────────────────────────────────────────

def test_read_offerte_reads_item_rows_and_skips_headers():
    header = [None] * 22
    header[1] = "Art. nr.  Jumbo"
    header[5] = "Omschrijving Item"
    item = [None] * 22
    item[1] = "No. 1"
    item[4] = "Algemeen"
    item[5] = "KLIC-melding"
    item[7] = "Post"
    item[9] = 1
    item[21] = 69.34
    blank = [None] * 22  # no Groep -> not an item row

    xlsx = _xlsx_bytes("vertaling bidsheet - offerte", [header, item, blank])

    items = read_offerte(xlsx, sheet_name="vertaling bidsheet - offerte")

    assert len(items) == 1
    assert items[0].groep == "Algemeen"
    assert items[0].omschrijving == "KLIC-melding"
    assert items[0].aantal == 1.0
    assert items[0].totaal == 69.34


def test_read_offerte_missing_sheet_returns_empty():
    xlsx = _xlsx_bytes("Sheet1", [[1, 2, 3]])
    assert read_offerte(xlsx, sheet_name="does not exist") == []


# ─── whole-section IB total ──────────────────────────────────────────────

def test_get_sloopwerk_totaal_ib_sums_subheadings():
    bdf = _budget_df({
        18: ("A.01.01", "Algemeen", 0, None, 3899.16),
        28: ("A.01.02", "Reiskosten", 0, None, 9037.0),
    })

    summed = get_sloopwerk_totaal_ib(bdf, row_start=18, row_end=28)

    assert summed == pytest.approx(3899.16 + 9037.0)


def test_get_sloopwerk_totaal_ib_none_bdf_returns_none():
    assert get_sloopwerk_totaal_ib(None) is None


# ─── matching ──────────────────────────────────────────────────────────────

def test_match_sloopwerk_budget_none_bdf_returns_empty():
    assert match_sloopwerk_budget([_offerte_item()], None, {}) == {}


def test_match_sloopwerk_budget_skips_nrfp_rows():
    offerte = [_offerte_item(groep="NRFP", omschrijving="Precariokosten")]
    bdf = _budget_df({19: ("A.01.01.01", "Precariokosten", 0, 100.0, 0.0)})
    mapping = {"Algemeen": [("A.01.01", "Algemeen")]}

    matches = match_sloopwerk_budget(offerte, bdf, mapping, row_start=19, row_end=19)

    assert matches == {}


def test_match_sloopwerk_budget_skips_groep_without_categorie_mapping():
    # 'Additionele kosten' has no entry in categorie_mapping -> left
    # unmatched rather than fuzzy-matched against the whole section.
    offerte = [_offerte_item(groep="Additionele kosten", omschrijving="KLIC-melding")]
    bdf = _budget_df({19: ("A.01.01.02", "KLIC-melding", 1, 69.34, 69.34)})
    mapping = {"Algemeen": [("A.01.01", "Algemeen")]}

    matches = match_sloopwerk_budget(offerte, bdf, mapping, row_start=19, row_end=19)

    assert matches == {}


@pytest.mark.skipif(not HAS_RAPIDFUZZ, reason="rapidfuzz not installed")
def test_match_sloopwerk_budget_restricts_to_same_categorie():
    # Two Budget rows with near-identical text live under different
    # categories — the offerte item's own Groep must pick the one in its
    # own categorie, not the higher-scoring one in an unrelated category.
    offerte = [_offerte_item(groep="Vergunningen", omschrijving="Afstemming Gemeente")]
    bdf = _budget_df({
        19: ("A.01.01.02", "Afstemming Gemeente (verkeerd genest)", 0, 100.0, 0.0),
        39: ("A.01.03.03", "Afstemming Gemeente", 1, 341.61, 341.61),
    })
    mapping = {
        "Algemeen": [("A.01.01", "Algemeen")],
        "Vergunningen": [("A.01.03", "Vergunningen")],
    }

    matches = match_sloopwerk_budget(offerte, bdf, mapping, row_start=19, row_end=39)

    assert matches[0].budget_row == 39


# ─── df building / status ───────────────────────────────────────────────────

def test_build_sloopwerk_df_flags_price_and_quantity_differences():
    offerte = [
        _offerte_item(row=1, totaal=100.0, aantal=2),   # price mismatch
        _offerte_item(row=2, totaal=200.0, aantal=3),   # qty mismatch
        _offerte_item(row=3, totaal=180.0, aantal=2),   # ok
        _offerte_item(row=4, groep="Vergunningen"),      # unmatched
        _offerte_item(row=5, groep="NRFP"),               # nrfp
    ]
    match_row = BudgetMatch(budget_row=19, code="A.01.01.01", omschrijving="x",
                             aantal=2, prijs=90.0, totaal=180.0, score=100.0, source="fuzzy")
    matches = {
        0: match_row,
        1: BudgetMatch(budget_row=20, code="A.01.01.02", omschrijving="x",
                       aantal=2, prijs=100.0, totaal=200.0, score=100.0, source="fuzzy"),
        2: BudgetMatch(budget_row=21, code="A.01.01.03", omschrijving="x",
                       aantal=2, prijs=90.0, totaal=180.0, score=100.0, source="fuzzy"),
    }

    df = build_sloopwerk_df(offerte, matches)

    assert list(df["_status"]) == ["prijs_afwijking", "aantal_afwijking", "ok", "unmatched", "nrfp"]


def test_build_sloopwerk_df_none_match_is_unmatched():
    offerte = [_offerte_item()]

    df = build_sloopwerk_df(offerte, {})

    assert df.loc[0, "_status"] == "unmatched"
    assert df.loc[0, "IB omschrijving"] is None


# ─── shared 'Alle resultaten' comparison ─────────────────────────────────────

def test_build_sloopwerk_vergelijking_derives_prijs_from_totaal():
    offerte = [
        _offerte_item(row=1, totaal=150.0, aantal=3),
        _offerte_item(row=2, groep="Vergunningen", omschrijving="Sloopmelding", totaal=None, aantal=1),
        _offerte_item(row=3, groep="NRFP", totaal=50.0, aantal=1),
    ]
    matches = {0: BudgetMatch(budget_row=19, code="A.01.01.01", omschrijving="KLIC", aantal=3, prijs=40.0,
                              totaal=120.0, score=88.0, source="fuzzy")}

    v = build_sloopwerk_vergelijking(offerte, matches)

    assert v.namen == ["Van de Laar"]
    assert v.categorieen == ["Algemeen", "Vergunningen"]
    klic, melding = v.regels
    assert klic.offertes == [(3, 50.0)] and klic.offerte_totaal(0) == pytest.approx(150.0)
    assert klic.ib_totaal == 120.0 and klic.opmerking == "fuzzy match 88%"
    assert melding.offertes == [(1, None)] and melding.offerte_totaal(0) == 0
    assert melding.opmerking == "geen IB-match"
