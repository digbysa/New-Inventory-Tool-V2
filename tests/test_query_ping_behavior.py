from pathlib import Path


SCRIPT_PATH = Path(__file__).resolve().parents[1] / "NewAssetTool.Wpf.matching.v14.ps1"
SCRIPT = SCRIPT_PATH.read_text(encoding="utf-8-sig")


def test_query_schedules_ping_button_after_two_seconds():
    assert "function Start-DelayedQueryPing" in SCRIPT
    assert "[int]$DelayMilliseconds=2000" in SCRIPT
    assert "Start-DelayedQueryPing -Ui $ui -QueryToken $script:AppState.CurrentQueryToken" in SCRIPT
    assert "$state.Ui.PingButton.RaiseEvent" in SCRIPT


def test_only_explicit_ping_click_starts_continuous_cmd_ping():
    assert "if ($script:AppState.AutomatedPingClick)" in SCRIPT
    automated_branch, explicit_branch = SCRIPT.split(
        "if ($script:AppState.AutomatedPingClick)", 1
    )[1].split("else", 1)
    assert "-StartContinuous" not in automated_branch
    assert "-StartContinuous" in explicit_branch


def test_nearby_ping_only_reports_network_details_after_icmp_success():
    nearby_ping = SCRIPT.split("function Start-NearbyRowsPingAsync", 1)[1].split(
        "function Invoke-SelectedNearbyPing", 1
    )[0]

    assert "DNS can retain an address for an offline or reassigned host" in nearby_ping
    assert nearby_ping.count("[pscustomobject]@{ IpAddress=''; Subnet=''; Success=$false }") == 3
    assert "Success=(-not [string]::IsNullOrWhiteSpace($ipAddress))" not in nearby_ping
    assert "$resolved = $Result -and $Result.Success" in SCRIPT
    assert "$Row.IPAddress = if ($resolved)" in SCRIPT
    assert "$Row.Subnet = if ($resolved -and $Result.Subnet)" in SCRIPT


def test_nearby_ping_progress_owns_status_badge_until_completion():
    assert "NearbyPingInProgress=$false" in SCRIPT
    assert "$script:AppState.NearbyPingInProgress = $true" in SCRIPT
    assert "$script:AppState.NearbyPingInProgress = $false" in SCRIPT
    assert "$Mode -notin @('Pinging','PingComplete','Warning')" in SCRIPT
    assert "devices…" not in SCRIPT
