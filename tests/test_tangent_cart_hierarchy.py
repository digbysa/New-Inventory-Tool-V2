import re
from pathlib import Path


SCRIPT_PATH = Path(__file__).resolve().parents[1] / "NewAssetTool.Wpf.matching.v14.ps1"
SCRIPT = SCRIPT_PATH.read_text(encoding="utf-8-sig")


def function_body(name: str) -> str:
    match = re.search(rf"function {name} \{{(?P<body>.*?)\n    \}}", SCRIPT, re.DOTALL)
    assert match
    return match.group("body")


def test_tangent_peripherals_use_the_named_cart_as_their_parent():
    resolver = function_body("Get-PeripheralAssociationParent")
    dialog = SCRIPT.split("function Show-AddPeripheralDialog", 1)[1].split(
        "function Get-CmdbLink", 1
    )[0]

    assert "-notmatch '^(?i)AO'" in resolver
    assert '$cartName = "$($ParentDevice.Name)-CRT"' in resolver
    assert "$cart.DetectedType -eq 'Cart'" in resolver
    assert "-NewParent $associationParent.AssetTag" in dialog
    assert "'u_parent_asset' -Value $associationParent.AssetTag" in dialog


def test_cart_peripherals_are_listed_as_grandchildren():
    associated = function_body("Build-AssociatedDevices")
    parent_resolver = function_body("Resolve-ParentDevice")

    assert "$child.DetectedType -ne 'Cart'" in associated
    assert "Role='Grandchild'" in associated
    assert "$record.DetectedType -eq 'Cart'" in parent_resolver
    assert "$computer.DetectedType -eq 'Computer'" in parent_resolver
