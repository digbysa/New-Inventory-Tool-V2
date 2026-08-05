from pathlib import Path

SCRIPT = (Path(__file__).resolve().parents[1] / "NewAssetTool.Wpf.matching.v14.ps1").read_text()


def test_nearby_context_menu_places_isolate_above_ping_selected_hosts():
    context_menu = SCRIPT.split("function Initialize-NearbyContextMenu", 1)[1].split("$Ui.NearbyDataGrid.ContextMenu = $menu", 1)[0]
    assert "Header='Isolate'" in context_menu
    assert "Header='Ping selected host(s)'" in context_menu
    assert context_menu.index("Header='Isolate'") < context_menu.index("Header='Ping selected host(s)'")


def test_isolate_filters_nearby_rows_to_selected_host_names():
    assert "function Invoke-IsolateSelectedNearbyRows" in SCRIPT
    isolate_body = SCRIPT.split("function Invoke-IsolateSelectedNearbyRows", 1)[1].split("function Invoke-SelectedNearbyPing", 1)[0]
    assert "Get-NearbySelectedRows -Ui $Ui" in isolate_body
    assert "$script:AppState.NearbyIsolatedHostNames.Clear()" in isolate_body
    assert "$script:AppState.NearbyIsolatedHostNames.Add($hostKey)" in isolate_body
    visible_body = SCRIPT.split("function Test-NearbyRowVisible", 1)[1].split("function AutoFit-NearbyColumns", 1)[0]
    assert "NearbyIsolatedHostNames.Count -gt 0" in visible_body
    assert "-not $script:AppState.NearbyIsolatedHostNames.Contains($hostKey)" in visible_body


def test_show_all_clears_nearby_isolation_filter():
    show_all_hook = SCRIPT.split("$ui.ShowAllNearbyButton.Add_Click", 1)[1].split("$ui.ShowAllNearbyCheckBox.Add_Click", 1)[0]
    assert "$script:AppState.NearbyIsolatedHostNames.Clear()" in show_all_hook
