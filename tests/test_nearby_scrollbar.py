from pathlib import Path
import xml.etree.ElementTree as ET


ROOT = Path(__file__).resolve().parents[1]
XAML_PATH = ROOT / "NewAssetTool.matching.v14.xaml"
XAML = XAML_PATH.read_text(encoding="utf-8")


def test_nearby_scrollbar_sizes_each_orientation_without_collapsing_horizontal_track():
    nearby_grid = XAML.split('<DataGrid x:Name="NearbyDataGrid"', 1)[1].split(
        "</DataGrid>", 1
    )[0]
    scrollbar_style = nearby_grid.split(
        '<Style TargetType="{x:Type ScrollBar}">', 1
    )[1].split("</Style>", 1)[0]

    assert '<Trigger Property="Orientation" Value="Vertical">' in scrollbar_style
    assert '<Setter Property="Width" Value="13"/>' in scrollbar_style
    assert '<Trigger Property="Orientation" Value="Horizontal">' in scrollbar_style
    assert '<Setter Property="Height" Value="13"/>' in scrollbar_style
    assert scrollbar_style.index('Value="Vertical"') < scrollbar_style.index(
        'Property="Width"'
    )
    assert scrollbar_style.index('Value="Horizontal"') < scrollbar_style.index(
        'Property="Height"'
    )


def test_xaml_remains_well_formed():
    ET.parse(XAML_PATH)
