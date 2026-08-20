import io

import openpyxl
import pytest

from matching_van_keulen import (
    IBItem,
    NettoItem,
    PdfItem,
    HAS_RAPIDFUZZ,
    _dutch,
    _skip,
    apply_art_nr_mapping,
    build_matches,
    read_budget_summary_rows,
    read_ib_items,
    results_to_df,
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


# ─── small parsing helpers ──────────────────────────────────────────────────

@pytest.mark.parametrize("raw,expected", [
    ("1.234,56", 1234.56),
    ("12,00", 12.0),
])
def test_dutch_number_format(raw, expected):
    assert _dutch(raw) == expected


def test_skip_filters_boilerplate_lines():
    assert _skip("") is True
    assert _skip("Van Keulen BV") is True
    assert _skip("1234 AB Amsterdam") is True
    assert _skip("Een normale offerteregel") is False


# ─── art.nr. mapping (Budget column BA backfill) ────────────────────────────

def test_apply_art_nr_mapping_single_candidate_always_applied():
    ii = _ib_item(description="Deurpaneel wit", pdf_art_nr="")
    mapping = {"deurpaneel wit": ["12345"]}

    filled = apply_art_nr_mapping([ii], mapping)

    assert filled == 1
    assert ii.pdf_art_nr == "12345"


def test_apply_art_nr_mapping_multi_candidate_resolved_by_pdf_art_nrs():
    # 'wandstelling...' is reused across 3 real Budget rows, each a
    # different size variant with its own Art.nr. — only the one that
    # actually appears in this offerte's own PDF should be applied.
    ii = _ib_item(description="Wandstelling 1400mm, voet 560mm, schappen 560mm", pdf_art_nr="")
    mapping = {"wandstelling 1400mm voet 560mm schappen 560mm": ["279738", "279739", "279740"]}

    filled = apply_art_nr_mapping([ii], mapping, pdf_art_nrs={"279739"})

    assert filled == 1
    assert ii.pdf_art_nr == "279739"


def test_apply_art_nr_mapping_multi_candidate_left_unresolved_when_ambiguous():
    ii = _ib_item(description="Wandstelling 1400mm, voet 560mm, schappen 560mm", pdf_art_nr="")
    mapping = {"wandstelling 1400mm voet 560mm schappen 560mm": ["279738", "279739", "279740"]}

    # None of the offerte's own art.nrs are among the candidates.
    filled = apply_art_nr_mapping([ii], mapping, pdf_art_nrs={"999999"})
    assert filled == 0
    assert ii.pdf_art_nr == ""

    # More than one candidate present — still ambiguous, can't tell which
    # Budget row this is.
    filled = apply_art_nr_mapping([ii], mapping, pdf_art_nrs={"279738", "279740"})
    assert filled == 0
    assert ii.pdf_art_nr == ""


def test_apply_art_nr_mapping_skips_rows_with_existing_pdf_art_nr():
    ii = _ib_item(description="Deurpaneel wit", pdf_art_nr="already-set")
    mapping = {"deurpaneel wit": ["12345"]}

    filled = apply_art_nr_mapping([ii], mapping)

    assert filled == 0
    assert ii.pdf_art_nr == "already-set"


# ─── matching ────────────────────────────────────────────────────────────────

def _pdf_item(**overrides) -> PdfItem:
    base = dict(art_nr="12345", description="Deurpaneel wit", quantity=2.0, unit_price=100.0,
                total=200.0, section="Deuren", manual_ref="", page=2)
    base.update(overrides)
    return PdfItem(**base)


def _netto_item(**overrides) -> NettoItem:
    base = dict(art_nr_jumbo="JUM-1", manual_nr="", art_nr_leverancier="12345", groep="Deuren",
                description="Deurpaneel wit", netto_price=90.0, unit="stuks")
    base.update(overrides)
    return NettoItem(**base)


def _ib_item(**overrides) -> IBItem:
    base = dict(nummer="6", description="Deurpaneel wit", manual="", unit="stuks",
                quantity=2.0, price=90.0, pdf_art_nr="12345", row=1440)
    base.update(overrides)
    return IBItem(**base)


def test_build_matches_exact_artnr():
    results = build_matches([_pdf_item()], [_netto_item()], [_ib_item()])

    assert len(results) == 1
    mr = results[0]
    assert mr.pdf_match_method == "exact_artnr"
    assert mr.ib_match_method == "exact_artnr"
    assert mr.confidence == 1.0


def test_build_matches_netto_artnr_can_list_multiple_codes():
    netto = _netto_item(art_nr_leverancier="99999+12345")
    results = build_matches([_pdf_item()], [netto], [])

    assert results[0].netto_item is netto
    assert results[0].pdf_match_method == "exact_artnr"


def test_build_matches_manual_code_single_candidate():
    pdf_item = _pdf_item(art_nr="00000", manual_ref="ABC")
    netto = _netto_item(art_nr_leverancier="", manual_nr="ABC")

    results = build_matches([pdf_item], [netto], [])

    assert results[0].netto_item is netto
    assert results[0].pdf_match_method == "manual_code"
    assert results[0].confidence == 0.85


def test_build_matches_ib_row_override_bypasses_normal_lookup():
    pdf_item = _pdf_item(art_nr="334363")
    override_item = _ib_item(row=2612, pdf_art_nr="334363")

    results = build_matches([pdf_item], [], [], ib_row_override={"334363": override_item})

    assert results[0].ib_match_method == "manual_row"
    assert results[0].ib_items == [override_item]


def test_build_matches_unmatched_when_nothing_lines_up():
    pdf_item = _pdf_item(art_nr="00000", manual_ref="", description="Totally unrelated widget")

    results = build_matches([pdf_item], [], [], fuzzy_threshold=100)

    assert results[0].pdf_match_method == "unmatched"
    assert results[0].ib_match_method == "unmatched"


@pytest.mark.skipif(not HAS_RAPIDFUZZ, reason="rapidfuzz not installed")
def test_build_matches_fuzzy_description_fallback():
    pdf_item = _pdf_item(art_nr="00000", manual_ref="", description="Deurpaneel wit RAL9010")
    netto = _netto_item(art_nr_leverancier="", manual_nr="")

    results = build_matches([pdf_item], [netto], [], fuzzy_threshold=50)

    assert results[0].pdf_match_method == "fuzzy"
    assert results[0].netto_item is netto


# ─── results -> dataframe ───────────────────────────────────────────────────

def test_results_to_df_price_diff_only_on_exact_match():
    results = build_matches([_pdf_item(unit_price=100.0)], [_netto_item(netto_price=90.0)], [_ib_item()])

    df = results_to_df(results)

    assert df.loc[0, "_status"] == "ok"
    assert df.loc[0, "Prijs verschil"].startswith("↑")


def test_results_to_df_flags_quantity_difference():
    pdf_item = _pdf_item(quantity=5.0)
    ib_item = _ib_item(quantity=2.0)
    results = build_matches([pdf_item], [_netto_item()], [ib_item])

    df = results_to_df(results)

    assert df.loc[0, "_status"] == "qty_diff"
    assert df.loc[0, "Aantal verschil"] == 3.0


def test_results_to_df_groups_duplicate_art_nr_and_sums_quantity():
    results = build_matches(
        [_pdf_item(quantity=2.0, total=200.0), _pdf_item(quantity=3.0, total=300.0)],
        [_netto_item()], [_ib_item(quantity=5.0)],
    )

    df = results_to_df(results)

    assert len(df) == 1
    assert df.loc[0, "PDF Aantal"] == 5.0
    assert df.loc[0, "_status"] == "ok"


# ─── budget sheet readers ───────────────────────────────────────────────────

def test_read_ib_items_filters_by_column_f_and_reads_pdf_art_nr():
    header = [None] * 60
    # Non-null text in columns F and BA keeps those columns dtype=object, so the
    # numeric-looking strings below survive as strings instead of being coerced
    # to float by pandas (which would happen if the column were all-NaN + digits).
    header[5] = "Header F"
    header[52] = "Header BA"
    row_ok = [None] * 60
    row_ok[5] = "2"
    row_ok[7] = "Deurpaneel wit"
    row_ok[10] = 2
    row_ok[11] = 90.0
    row_ok[52] = "12345"
    row_skip = [None] * 60
    row_skip[5] = "1"  # not "2" -> excluded
    row_skip[7] = "Irrelevant"

    xlsx = _xlsx_bytes("Budget", [header, row_ok, row_skip])

    items = read_ib_items(io.BytesIO(xlsx), row_start=2, row_end=3)

    assert len(items) == 1
    assert items[0].description == "Deurpaneel wit"
    assert items[0].pdf_art_nr == "12345"
    assert items[0].quantity == 2.0
    assert items[0].price == 90.0


def test_read_budget_summary_rows_excludes_lamellen_plafond_when_qty_over_one():
    def _row(nummer, naam, aantal, prijs):
        row = [None] * 12
        row[6], row[7], row[9], row[10], row[11] = nummer, naam, "stuks", aantal, prijs
        return row

    rows = {
        2263: _row("1", "Lamellen Plafond", 3, 50.0),  # excluded: aantal > 1
        2264: _row("2", "Lamellen Plafond", 1, 50.0),  # kept: aantal == 1
        2265: _row("3", "Gewoon item", 2, 25.0),
    }
    sheet_rows = [[None] * 12 for _ in range(max(rows))]
    for r, values in rows.items():
        sheet_rows[r - 1] = values

    xlsx = _xlsx_bytes("Budget", sheet_rows)

    kept, excluded = read_budget_summary_rows(io.BytesIO(xlsx), 2263, 2265)

    assert [r["rij"] for r in excluded] == [2263]
    assert {r["rij"] for r in kept} == {2264, 2265}
