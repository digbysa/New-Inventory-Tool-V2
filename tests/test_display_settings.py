from pathlib import Path
import xml.etree.ElementTree as ET

ROOT = Path(__file__).resolve().parents[1]
NS = '{http://schemas.microsoft.com/winfx/2006/xaml/presentation}'
XNAME = '{http://schemas.microsoft.com/winfx/2006/xaml}Name'
SCRIPT = (ROOT / 'NewAssetTool.Wpf.matching.v14.ps1').read_text(encoding='utf-8-sig')


def test_settings_gear_is_at_right_of_toolbar_and_registered():
    root = ET.parse(ROOT / 'NewAssetTool.matching.v14.xaml').getroot()
    gear = next(e for e in root.iter() if e.get(XNAME) == 'SettingsButton')
    assert gear.get('Grid.Column') == '7'
    assert gear.get('HorizontalAlignment') == 'Right'
    assert 'Display settings' == gear.get('AutomationProperties.Name')
    assert "'SettingsButton'" in SCRIPT.split('$ui = Get-NamedControls', 1)[1].split('$ui.Window', 1)[0]
    assert '$ui.SettingsButton.Add_Click({ Show-DisplaySettings -Ui $ui })' in SCRIPT


def test_settings_dialog_xaml_has_scale_control_and_preview_save_cancel_reset():
    body = SCRIPT.split('function Show-DisplaySettings', 1)[1].split('function Update-SystemPageFit', 1)[0]
    markup = body.split("::Parse(@'", 1)[1].split("'@)", 1)[0]
    root = ET.fromstring(markup)
    slider = next(e for e in root.iter() if e.get(XNAME) == 'ScaleSlider')
    assert slider.get('Minimum') == '100'
    assert slider.get('Maximum') == '150'
    assert slider.get('IsSnapToTickEnabled') == 'True'
    assert any(e.get('IsCancel') == 'True' for e in root.iter(f'{NS}Button'))
    assert 'Set-SystemDisplayScale -Ui $sender.Tag.Ui' in body
    assert 'Save-DisplayScale -Percent' in body
    assert 'Set-SystemDisplayScale -Ui $Ui -Percent $previous' in body
    assert '$sender.Tag.Value = 100' in body
