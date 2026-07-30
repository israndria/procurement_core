import pytest

from inject_pl import MOD_NAME, _validate_vba_source


def test_validate_vba_source_accepts_vba_double_quotes():
    source = (
        f'Attribute VB_Name = "{MOD_NAME}"\n'
        'Public Sub Probe()\n'
        '    Range("A1").Formula = "=IF(B1="""","""",B1)"\n'
        "End Sub\n"
    )

    _validate_vba_source(source)


def test_validate_vba_source_rejects_python_style_formula_escape():
    source = (
        f'Attribute VB_Name = "{MOD_NAME}"\n'
        'Public Sub Probe()\n'
        '    Range("A1").Formula = "=IF(B1=\\"\\",B1)"\n'
        "End Sub\n"
    )

    with pytest.raises(ValueError, match="escape Python/JSON"):
        _validate_vba_source(source)
