import re
from pathlib import Path


ROOT = Path(__file__).parents[1]
SCRIPT = (ROOT / "NewAssetTool.Wpf.matching.v14.ps1").read_text(encoding="utf-8-sig")
XAML = (ROOT / "NewAssetTool.matching.v14.xaml").read_text(encoding="utf-8-sig")


def _function_body(name: str) -> str:
    match = re.search(
        rf"function {re.escape(name)} \{{(?P<body>.*?)\n    \}}", SCRIPT, re.DOTALL
    )
    assert match, f"Could not find {name}"
    return match.group("body")


def test_uptime_and_pending_reboot_follow_subnet_in_system_header():
    subnet = XAML.index('x:Name="DeviceSubnetText"')
    uptime = XAML.index('x:Name="DeviceUptimeText"')
    reboot = XAML.index('x:Name="DevicePendingRebootText"')

    assert subnet < uptime < reboot
    assert 'Text="Uptime: Unknown"' in XAML
    assert 'Text="Pending reboot: Unknown"' in XAML


def test_system_details_only_become_visible_after_successful_ping():
    visibility = _function_body("Set-DeviceNetworkVisibility")
    status_ui = _function_body("Set-OnlineStatusUi")
    ping = _function_body("Invoke-CurrentDevicePing")

    assert "$Ui.DeviceUptimeText.Visibility = $visibility" in visibility
    assert "$Ui.DevicePendingRebootText.Visibility = $visibility" in visibility
    assert "Set-DeviceNetworkVisibility -Ui $Ui -IsVisible:$IsOnline" in status_ui
    assert "if ($connectivity.IsOnline) { Get-RemoteRestartStatus" in ping
    assert "-Uptime $restartStatus.Uptime -PendingReboot $restartStatus.PendingReboot" in ping


def test_remote_restart_status_uses_short_dcom_cim_query():
    body = _function_body("Get-RemoteRestartStatus")

    assert "New-CimSessionOption -Protocol Dcom" in body
    assert "-OperationTimeoutSec 3" in body
    assert "-ClassName Win32_OperatingSystem" in body
    assert "LastBootUpTime" in body
    assert "Component Based Servicing\\RebootPending" in body
    assert "WindowsUpdate\\Auto Update\\RebootRequired" in body
    assert "Remove-CimSession" in body
