from vergelijking import leverancier_uit_rijen


def test_leverancier_uit_rijen_label_in_first_or_second_column():
    # Aannemer 'vertaling bidsheet': label in col 0
    rows = [(1, ["Algemene informatie"]), (2, ["Prijsopgave Leverancier", "Van de Laar", "NRFP"])]
    assert leverancier_uit_rijen(rows) == "Van de Laar"
    # Sloopwerk template: a numbering column first, label in col 1, empty cells before the name
    rows = [(9, [1.0, "Prijsopgave Leverancier", None, "  Van de Laar Sloopwerken "])]
    assert leverancier_uit_rijen(rows) == "Van de Laar Sloopwerken"


def test_leverancier_uit_rijen_none_without_label_or_beyond_max_rij():
    assert leverancier_uit_rijen([(1, ["Leverancier", "Frimex"])]) is None
    assert leverancier_uit_rijen([(41, ["Prijsopgave Leverancier", "Van Wijnen"])]) is None
