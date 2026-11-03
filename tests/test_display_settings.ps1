# Isolated PowerShell checks; fake value types let sizing run without WPF.
$ErrorActionPreference='Stop'
Set-StrictMode -Version 3
Add-Type @'
namespace System.Windows {
    public struct Thickness {
        public double Left,Top,Right,Bottom;
        public Thickness(double l,double t,double r,double b) { Left=l; Top=t; Right=r; Bottom=b; }
    }
    public class GridLength {
        public double Value;
        public bool IsAbsolute;
        public GridLength(double v) { Value=v; IsAbsolute=true; }
        public static GridLength Auto { get { return new GridLength(0) { IsAbsolute=false }; } }
    }
}
namespace System.Windows.Threading {
    public enum DispatcherPriority { Loaded }
    public class FakeDispatcher {
        public object BeginInvoke(System.Action a, DispatcherPriority p) { return null; }
    }
}
'@
$tokens=$null; $parseErrors=$null
$ast=[System.Management.Automation.Language.Parser]::ParseFile((Join-Path $PSScriptRoot '../NewAssetTool.Wpf.matching.v14.ps1'),[ref]$tokens,[ref]$parseErrors)
if ($parseErrors.Count) { throw ($parseErrors | Out-String) }
foreach ($name in @('Read-DisplayScale','Save-DisplayScale','Set-SystemDisplayScale')) {
    $node=$ast.Find({param($n) $n -is [System.Management.Automation.Language.FunctionDefinitionAst] -and $n.Name -eq $name},$true)
    Invoke-Expression $node.Extent.Text
}
$folder=Join-Path ([IO.Path]::GetTempPath()) ('display-settings-test-'+[guid]::NewGuid().ToString('N'))
$path=Join-Path $folder 'settings.json'
try {
    if ((Read-DisplayScale -Path $path) -ne 100) { throw 'Missing settings should default to 100.' }
    Save-DisplayScale -Percent 125 -Path $path
    if ((Read-DisplayScale -Path $path) -ne 125) { throw 'Saved scale did not round-trip.' }
    Save-DisplayScale -Percent 999 -Path $path
    if ((Read-DisplayScale -Path $path) -ne 150) { throw 'Saved scale exceeded maximum.' }
    Set-Content -LiteralPath $path -Value 'invalid JSON'
    if ((Read-DisplayScale -Path $path) -ne 100) { throw 'Corrupt settings should fall back to 100.' }
} finally { if (Test-Path -LiteralPath $folder) { Remove-Item -LiteralPath $folder -Recurse -Force } }
$margin=[System.Windows.Thickness]::new(12,10,12,10)
$element=[pscustomobject]@{ FontSize=14.0; Height=400.0; MinHeight=350.0; Margin=$margin }
$row=[pscustomobject]@{ Height=[System.Windows.GridLength]::new(48) }
$column=[pscustomobject]@{ Width=[System.Windows.GridLength]::new(95) }
$script:SystemDisplayEntries=@([pscustomobject]@{
    Element=$element; FontSize=14.0; Height=400.0; MinHeight=350.0; Margin=$margin
    Rows=@([pscustomobject]@{ Row=$row; Height=$row.Height })
    Columns=@([pscustomobject]@{ Column=$column; Width=$column.Width })
})
$script:SystemDisplayScale=100
$ui=@{ CommentsTextBox=[pscustomobject]@{ MinHeight=0 }; SystemPageHost=[pscustomobject]@{ Dispatcher=[System.Windows.Threading.FakeDispatcher]::new() } }
Set-SystemDisplayScale -Ui $ui -Percent 125
if ($element.FontSize -ne 17.5 -or $element.MinHeight -ne 0 -or -not [double]::IsNaN($element.Height)) { throw 'Larger mode did not enlarge text and reclaim fixed space.' }
if ($row.Height.IsAbsolute) { throw 'Headers should follow their contents.' }
Set-SystemDisplayScale -Ui $ui -Percent 150
if ($element.FontSize -ne 21) { throw 'Scale compounded instead of using baseline.' }
Set-SystemDisplayScale -Ui $ui -Percent 100
if ($element.FontSize -ne 14 -or $element.Height -ne 400 -or $element.MinHeight -ne 350 -or $element.Margin.Top -ne 10 -or $row.Height.Value -ne 48 -or $column.Width.Value -ne 95) { throw 'Reset did not restore baseline geometry.' }
'PASS: preference persistence, bounds, corrupt-file fallback, larger fonts, compact spacing, noncompounding preview, reset.'
