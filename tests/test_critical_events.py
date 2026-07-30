import re
import xml.etree.ElementTree as ET
from pathlib import Path


ROOT = Path(__file__).resolve().parents[1]
SCRIPT = (ROOT / "NewAssetTool.Wpf.matching.v14.ps1").read_text(encoding="utf-8")
XAML = (ROOT / "NewAssetTool.matching.v14.xaml").read_text(encoding="utf-8")


def test_dialog_xaml_declares_the_x_namespace_and_is_valid_xml():
    function = SCRIPT.split("function Show-CriticalEventsDialog", 1)[1].split(
        "function Invoke-CurrentDevicePing", 1
    )[0]
    dialog_xaml = re.search(r"\[xml\]\$dialogXaml = @'\s*(.*?)\s*'@", function, re.DOTALL)

    assert dialog_xaml, "Critical Events dialog XAML must be present"
    root = ET.fromstring(dialog_xaml.group(1))
    assert root.tag == "{http://schemas.microsoft.com/winfx/2006/xaml/presentation}Window"
    assert 'x:Name="TitleText"' in dialog_xaml.group(1)


def test_button_is_immediately_right_of_lookup_subnet():
    lookup = XAML.index('x:Name="LookupSubnetButton"')
    critical = XAML.index('x:Name="CriticalEventsButton"')
    assert lookup < critical
    assert 'Grid.Column="6"' in XAML[critical : critical + 180]


def test_button_is_registered_before_click_handler_is_attached():
    controls = SCRIPT.split("$ui = Get-NamedControls", 1)[1].split("$ui.Window = $window", 1)[0]
    assert "'CriticalEventsButton'" in controls
    assert SCRIPT.index("'CriticalEventsButton'", SCRIPT.index("$ui = Get-NamedControls")) < SCRIPT.index(
        "$ui.CriticalEventsButton.Add_Click"
    )


def test_remote_query_filters_system_log_at_source():
    query = SCRIPT.split("function Show-CriticalEventsDialog", 1)[1].split("function Invoke-CurrentDevicePing", 1)[0]
    assert "Get-WinEvent -ComputerName $TargetComputer -FilterHashtable" in query
    assert "LogName='System'" in query
    assert "StartTime=$StartTime" in query
    assert "Id=41,86,88,1001,6008" in query
    assert "Sort-Object TimeCreated -Descending" in query


def test_remote_query_falls_back_to_wmi_over_dcom():
    query = SCRIPT.split("function Show-CriticalEventsDialog", 1)[1].split("function Invoke-CurrentDevicePing", 1)[0]
    assert "New-CimSessionOption -Protocol Dcom" in query
    assert "Get-CimInstance -CimSession $cimSession -ClassName Win32_NTLogEvent" in query
    assert "LogFile='System' AND TimeGenerated >= '$dmtfStart'" in query
    assert "QueryMethod='WMI (DCOM) fallback'" in query
    assert "Remove-CimSession" in query


def test_outcomes_and_async_timeout_are_present():
    assert "BeginInvoke()" in SCRIPT
    assert "The PowerShell command timed out" in SCRIPT
    assert "No crashes or unexpected shutdowns were recorded in the last 7 days." in SCRIPT
    assert "Total matching events:" in SCRIPT
    assert "The computer could not be queried." in SCRIPT
    assert "malformed or incomplete event data" in SCRIPT
    assert "Both Event Log RPC and the WMI fallback failed." in SCRIPT


def test_event_rows_sort_and_explain_shutdown_evidence_and_bugchecks():
    converter = SCRIPT.split("function ConvertTo-CriticalEventRows", 1)[1].split("function Show-CriticalEventsDialog", 1)[0]
    interpretation = SCRIPT.split("function Get-CriticalEventInterpretation", 1)[1].split("function ConvertTo-CriticalEventRows", 1)[0]
    assert "Sort-Object TimeCreated -Descending" in converter
    assert "TotalMinutes) -le 15" in converter
    assert "does not identify the cause" in interpretation
    assert "possible blue-screen crash" in interpretation
    for event_id in (41, 86, 88, 1001, 6008):
        assert str(event_id) in interpretation
