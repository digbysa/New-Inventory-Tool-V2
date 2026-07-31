from pathlib import Path


ROOT = Path(__file__).resolve().parents[1]
XAML = (ROOT / "NewAssetTool.matching.v14.xaml").read_text(encoding="utf-8")
SCRIPT = (ROOT / "NewAssetTool.Wpf.matching.v14.ps1").read_text(encoding="utf-8")


def test_disabled_manual_round_button_uses_grey_style_and_shows_tooltip():
    assert 'x:Key="ManualRoundButtonStyle"' in XAML
    assert '<Setter Property="ToolTipService.ShowOnDisabled" Value="True"/>' in XAML
    assert '<Setter Property="Background" Value="#94A3B8"/>' in XAML
    assert 'Style="{StaticResource ManualRoundButtonStyle}"' in XAML


def test_manual_round_disabled_states_explain_why_the_action_is_unavailable():
    assert "Manual Round is unavailable until a device is selected." in SCRIPT
    assert "selected device has no asset tag." in SCRIPT
    assert "no rounding URL was found for asset tag" in SCRIPT
    assert "$Ui.ManualRoundButton.ToolTip = $reason" in SCRIPT


def test_failed_webpage_launch_disables_button_and_exposes_error_as_tooltip():
    assert '$ui.ManualRoundButton.IsEnabled = $false' in SCRIPT
    assert "rounding webpage could not be opened: $($_.Exception.Message)" in SCRIPT
