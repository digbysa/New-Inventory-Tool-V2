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


def test_isolate_uses_selection_captured_before_context_menu_takes_focus():
    isolate_body = SCRIPT.split("function Invoke-IsolateSelectedNearbyRows", 1)[1].split("function Invoke-SelectedNearbyPing", 1)[0]
    assert "$script:AppState.NearbyContextSelectedRows" in isolate_body
    assert "$selected = @($script:AppState.NearbyContextSelectedRows | Where-Object { $_ })" in isolate_body

    context_menu = SCRIPT.split("function Initialize-NearbyContextMenu", 1)[1].split("$Ui.NearbyDataGrid.Add_MouseDoubleClick", 1)[0]
    assert "$script:AppState.NearbyContextSelectedRows = @(Get-NearbySelectedRows -Ui $Ui)" in context_menu


def test_show_all_clears_nearby_isolation_filter():
    show_all_hook = SCRIPT.split("$ui.ShowAllNearbyButton.Add_Click", 1)[1].split("$ui.ShowAllNearbyCheckBox.Add_Click", 1)[0]
    assert "$script:AppState.NearbyIsolatedHostNames.Clear()" in show_all_hook


def test_isolate_handler_uses_main_script_state_instead_of_a_closure_module():
    handler = SCRIPT.split("$isolateItem.Add_Click(", 1)[1].split("[void]$menu.Items.Add($isolateItem)", 1)[0]
    assert ".GetNewClosure()" not in handler
    assert "-Inventory $script:AppState.Inventory" in handler
    assert "-Ui $sender.Tag.Ui" in handler


def test_restore_refreshes_and_returns_before_changing_filter_checkboxes():
    handler = SCRIPT.split("$ui.ShowAllNearbyButton.Add_Click", 1)[1].split("$ui.ShowAllNearbyCheckBox.Add_Click", 1)[0]
    restore, show_all = handler.split("return", 1)
    assert "NearbyIsolatedHostNames.Count -gt 0" in restore
    assert "NearbyIsolatedHostNames.Clear()" in restore
    assert "Update-NearbyRows -Ui $ui" in restore
    assert ".IsChecked =" not in restore
    assert "RaiseEvent" not in restore
    assert "$ui.ShowAllNearbyCheckBox.IsChecked = $true" in show_all


def test_nearby_summary_sets_button_label_from_isolation_state():
    summary = SCRIPT.split("function Update-NearbySummary", 1)[1].split("function Update-NearbyCheckboxLabels", 1)[0]
    assert "NearbyIsolatedHostNames.Count -gt 0" in summary
    assert "if ($isIsolated) { 'Restore' } else { 'Show All' }" in summary


def test_filter_only_refresh_reuses_rows_without_reloading_events():
    body = SCRIPT.split("function Update-NearbyRows", 1)[1].split("function Get-NearbySelectedRows", 1)[0]
    cached, rebuild = body.split("} else {", 1)
    assert "if ($FilterOnly -and $null -ne $script:NearbyAllRows)" in cached
    assert "$allRows = $script:NearbyAllRows" in cached
    assert "Load-NearbyRoundingEvents" not in cached
    assert "Build-NearbyDevices" not in cached
    assert "Build-NearbyDevices" in rebuild
    assert "$script:NearbyNaturalColumnWidths = $null" in rebuild
    isolate = SCRIPT.split("function Invoke-IsolateSelectedNearbyRows", 1)[1].split("function Invoke-SelectedNearbyPing", 1)[0]
    restore = SCRIPT.split("$ui.ShowAllNearbyButton.Add_Click", 1)[1].split("return", 1)[0]
    assert "-FilterOnly" in isolate
    assert "-FilterOnly" in restore
