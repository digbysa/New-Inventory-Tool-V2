import re
from pathlib import Path


SCRIPT_PATH = Path(__file__).parents[1] / "NewAssetTool.Wpf.matching.v14.ps1"


def test_first_location_hierarchy_row_is_not_rejected_as_an_empty_collection():
    script = SCRIPT_PATH.read_text(encoding="utf-8-sig")
    function = re.search(
        r"function Add-LocationHierarchyRow \{(?P<body>.*?)\n    \}",
        script,
        re.DOTALL,
    )

    assert function, "Add-LocationHierarchyRow must exist"
    body = function.group("body")
    assert "$null -eq $Rows" in body
    assert "if (-not $Rows" not in body
    assert "$Rows.Add(" in body
