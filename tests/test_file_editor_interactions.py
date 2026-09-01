import re
from pathlib import Path


SCRIPT_PATH = Path(__file__).parents[1] / "NewAssetTool.Wpf.matching.v14.ps1"


def _file_editor_body():
    script = SCRIPT_PATH.read_text(encoding="utf-8-sig")
    function = re.search(
        r"function Show-RoundingEventsFileEditor \{(?P<body>.*?)\n    \}",
        script,
        re.DOTALL,
    )
    assert function, "Show-RoundingEventsFileEditor must exist"
    return function.group("body")


def test_file_editor_supports_extended_full_row_selection_and_bulk_delete():
    body = _file_editor_body()

    assert "$grid.SelectionMode = 'Extended'" in body
    assert "$grid.SelectionUnit = 'FullRow'" in body
    assert "$grid.SelectedItems" in body
    assert "Header='Delete selected row(s)'" in body
    assert "foreach ($selectedRow in $selectedRows) { $selectedRow.Row.Delete() }" in body
    assert "$grid.Add_PreviewMouseRightButtonDown" in body
    assert "if (-not $sender.SelectedItems.Contains($clickedRow))" in body


def test_file_editor_supports_native_touchpad_horizontal_scrolling():
    body = _file_editor_body()

    assert "$editor.Add_SourceInitialized" in body
    assert "$msg -ne 0x020E -or -not $grid.IsMouseOver" in body
    assert "$source.AddHook($editorHorizontalWheelHook)" in body
    assert "$scrollViewer.LineRight()" in body
    assert "$scrollViewer.LineLeft()" in body
    assert "$handled.Value = $true" in body


def test_file_editor_supports_shift_wheel_horizontal_scrolling():
    body = _file_editor_body()

    assert "[System.Windows.Input.ModifierKeys]::Shift" in body
    assert "$grid.AddHandler([System.Windows.UIElement]::PreviewMouseWheelEvent" in body
