from types import SimpleNamespace

from word_merge import _resolve_stale_cached_formula_value


def _sheet(**values):
    return {key: SimpleNamespace(value=value) for key, value in values.items()}


def test_resolves_stale_cached_direct_reference_chain():
    formula_sheets = {
        "satu_data": _sheet(BO2="='@ Evaluasi'!C37"),
        "@ Evaluasi": _sheet(C37="='@ Master Data'!C61"),
        "@ Master Data": _sheet(C61="Nomor ND/100"),
    }
    cached_sheets = {
        "satu_data": _sheet(BO2="='@ Master Data'!C61"),
        "@ Evaluasi": _sheet(C37="='@ Master Data'!C61"),
        "@ Master Data": _sheet(C61="Nomor ND/100"),
    }

    resolved = _resolve_stale_cached_formula_value(
        cached_sheets["satu_data"]["BO2"].value,
        formula_sheets,
        cached_sheets,
        "satu_data",
        "BO2",
    )

    assert resolved == "Nomor ND/100"


def test_keeps_non_direct_formula_cache_untouched():
    formula_sheets = {"satu_data": _sheet(BO2='=TEXT(A2,"dd mmmm yyyy")')}
    cached_sheets = {"satu_data": _sheet(BO2="04 September 2026")}

    resolved = _resolve_stale_cached_formula_value(
        cached_sheets["satu_data"]["BO2"].value,
        formula_sheets,
        cached_sheets,
        "satu_data",
        "BO2",
    )

    assert resolved == "04 September 2026"
