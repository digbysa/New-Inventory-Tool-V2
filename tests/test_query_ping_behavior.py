from pathlib import Path


SCRIPT_PATH = Path(__file__).resolve().parents[1] / "NewAssetTool.Wpf.matching.v14.ps1"
SCRIPT = SCRIPT_PATH.read_text(encoding="utf-8-sig")


def test_query_schedules_ping_button_after_three_seconds():
    assert "function Start-DelayedQueryPing" in SCRIPT
    assert "[int]$DelayMilliseconds=3000" in SCRIPT
    assert "Start-DelayedQueryPing -Ui $ui -QueryToken $script:AppState.CurrentQueryToken" in SCRIPT
    assert "$state.Ui.PingButton.RaiseEvent" in SCRIPT


def test_only_explicit_ping_click_starts_continuous_cmd_ping():
    assert "if ($script:AppState.AutomatedPingClick)" in SCRIPT
    automated_branch, explicit_branch = SCRIPT.split(
        "if ($script:AppState.AutomatedPingClick)", 1
    )[1].split("else", 1)
    assert "-StartContinuous" not in automated_branch
    assert "-StartContinuous" in explicit_branch
