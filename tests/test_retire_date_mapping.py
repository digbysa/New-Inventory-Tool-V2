import re
from pathlib import Path


SCRIPT_PATH = Path(__file__).parents[1] / "NewAssetTool.Wpf.matching.v14.ps1"


def test_device_record_accepts_computer_and_peripheral_retire_date_fields():
    script = SCRIPT_PATH.read_text(encoding="utf-8-sig")
    match = re.search(
        r"RetireDate=Format-DateLong \(Get-FieldValue -Row \$Row -Names @\(([^)]*)\)\)",
        script,
    )

    assert match, "ConvertTo-DeviceRecord must map a retirement date"
    aliases = set(re.findall(r"'([^']+)'", match.group(1)))
    assert {"u_scheduled_retirement", "u_retired_date", "RetireDate"} <= aliases
