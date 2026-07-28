import re
from pathlib import Path


SCRIPT_PATH = Path(__file__).parents[1] / "NewAssetTool.Wpf.matching.v14.ps1"


def test_parent_and_child_association_indices_are_built_once():
    script = SCRIPT_PATH.read_text(encoding="utf-8-sig")
    build = re.search(
        r"function Build-InventoryIndices \{(?P<body>.*?)\n    \}", script, re.DOTALL
    )
    associated = re.search(
        r"function Build-AssociatedDevices \{(?P<body>.*?)\n    \}", script, re.DOTALL
    )

    assert build and associated
    assert "ChildrenByParent" in build.group("body")
    assert "$Inventory.ChildrenByParent" in associated.group("body")
    assert "foreach ($collectionName in @('Monitors','Carts','Mics','Scanners'))" not in associated.group("body")


def test_parent_resolution_uses_inventory_indices_instead_of_scanning_computers():
    script = SCRIPT_PATH.read_text(encoding="utf-8-sig")
    function = re.search(
        r"function Resolve-ParentDevice \{(?P<body>.*?)\n    \}", script, re.DOTALL
    )

    assert function
    body = function.group("body")
    assert "IndexByName" in body
    assert "IndexByAsset" in body
    assert "IndexBySerial" in body
    assert "foreach ($row in $Inventory.Computers)" not in body
