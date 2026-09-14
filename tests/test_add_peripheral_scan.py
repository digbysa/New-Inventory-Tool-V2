import re
from pathlib import Path


SCRIPT_PATH = Path(__file__).resolve().parents[1] / "NewAssetTool.Wpf.matching.v14.ps1"
SCRIPT = SCRIPT_PATH.read_text(encoding="utf-8-sig")


def test_add_peripheral_scan_removes_trailing_barcode_character_for_c0_tags():
    converter = re.search(
        r"function ConvertFrom-AddPeripheralScan \{(?P<body>.*?)\n    \}",
        SCRIPT,
        re.DOTALL,
    )

    assert converter
    body = converter.group("body")
    assert ".StartsWith('C0',[System.StringComparison]::OrdinalIgnoreCase)" in body
    assert ".Substring(0,$trimmed.Length - 1)" in body
    assert "return $Raw" in body


def test_add_peripheral_enter_uses_corrected_scan_for_lookup_and_display():
    dialog = SCRIPT.split("function Show-AddPeripheralDialog", 1)[1].split(
        "function Get-CmdbLink", 1
    )[0]

    assert "ConvertFrom-AddPeripheralScan -Raw $txt.Text" in dialog
    assert "$txt.Text = $query" in dialog
    assert "Resolve-AssociatedPeripheralLookup -Query $query" in dialog
    assert SCRIPT.count("ConvertFrom-AddPeripheralScan -Raw") == 1


def test_add_peripheral_dialog_focuses_search_when_displayed():
    dialog = SCRIPT.split("function Show-AddPeripheralDialog", 1)[1].split(
        "function Get-CmdbLink", 1
    )[0]

    assert "$window.Add_ContentRendered({" in dialog
    assert "$txt.Focus() | Out-Null" in dialog
    assert "$txt.CaretIndex = $txt.Text.Length" in dialog
    assert dialog.index("$window.Add_ContentRendered({") < dialog.index(
        "$window.ShowDialog()"
    )
