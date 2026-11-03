from pathlib import Path
import xml.etree.ElementTree as ET

ROOT = Path(__file__).resolve().parents[1]
NS = '{http://schemas.microsoft.com/winfx/2006/xaml/presentation}'
XNAME = '{http://schemas.microsoft.com/winfx/2006/xaml}Name'


def test_system_page_scales_down_to_fit_without_an_outer_scroll_viewer():
    root = ET.parse(ROOT / 'NewAssetTool.matching.v14.xaml').getroot()
    tab = next(e for e in root.iter() if e.get(XNAME) == 'SystemTab')
    viewport = tab.find(f'{NS}Viewbox')
    assert viewport is not None
    assert viewport.get('Stretch') == 'Uniform'
    assert viewport.get('StretchDirection') == 'DownOnly'
    assert tab.find(f'{NS}ScrollViewer') is None
    content = viewport.find(f'{NS}Grid')
    assert content.get('MinWidth') == '1200'
    assert 'SystemPageViewport' in content.get('Width')
    assert int(root.get('MinWidth')) <= 1000
    assert int(root.get('MinHeight')) <= 600
