from pathlib import Path


ROOT = Path(__file__).resolve().parents[1]
SCRIPT = (ROOT / "NewAssetTool.Wpf.matching.v14.ps1").read_text(encoding="utf-8")
XAML = (ROOT / "NewAssetTool.matching.v14.xaml").read_text(encoding="utf-8")


def test_lookup_subnet_button_is_to_the_right_of_monitor_label():
    monitor_position = XAML.index('x:Name="MonitorLabelButton"')
    lookup_position = XAML.index('x:Name="LookupSubnetButton"')

    assert monitor_position < lookup_position
    assert 'Content="Lookup Subnet"' in XAML
    assert 'Grid.Column="5"' in XAML[lookup_position : lookup_position + 180]


def test_lookup_dialog_validates_and_resolves_ipv4_input_from_data_root():
    dialog = SCRIPT.split("function Show-SubnetLookupDialog", 1)[1].split(
        "function Resolve-HostIPv4Address", 1
    )[0]

    assert "ConvertTo-IPv4Bytes -IpAddress $ipAddress" in dialog
    assert "Resolve-SubnetName -IpAddress $ipAddress -DataRoot $DataRoot" in dialog
    assert "Enter a valid IPv4 address." in dialog
    assert "No subnet was found for $ipAddress." in dialog
    assert '"Subnet: $subnetName"' in dialog


def test_lookup_button_opens_dialog_with_application_data_root():
    assert "'LookupSubnetButton'" in SCRIPT
    assert "$ui.LookupSubnetButton.Add_Click" in SCRIPT
    assert "Show-SubnetLookupDialog -Ui $ui -DataRoot $script:AppState.DataRoot" in SCRIPT
