$ErrorActionPreference='Stop'
$tokens=$null; $parseErrors=$null
$ast=[System.Management.Automation.Language.Parser]::ParseFile((Join-Path $PSScriptRoot '../NewAssetTool.Wpf.matching.v14.ps1'),[ref]$tokens,[ref]$parseErrors)
if ($parseErrors.Count) { throw ($parseErrors | Out-String) }
$node=$ast.Find({param($n) $n -is [System.Management.Automation.Language.FunctionDefinitionAst] -and $n.Name -eq 'Update-SystemPageFit'},$true)
Invoke-Expression $node.Extent.Text
foreach ($size in @(@(1870,900),@(1480,700),@(1250,560),@(950,440))) {
    $hostPanel=[pscustomobject]@{ IsVisible=$true; ActualWidth=$size[0]; ActualHeight=$size[1] }
    $content=[pscustomobject]@{ ActualHeight=850.0; MinWidth=1200.0; Width=1200.0 }
    Update-SystemPageFit -Ui @{ SystemPageHost=$hostPanel; SystemPageContent=$content }
    $scale=[Math]::Min(1.0,[Math]::Min($hostPanel.ActualHeight/$content.ActualHeight,$hostPanel.ActualWidth/$content.Width))
    if ([Math]::Abs($content.Width*$scale-$hostPanel.ActualWidth) -gt 0.5) { throw "Page did not fill width at $size" }
    if ($content.ActualHeight*$scale -gt $hostPanel.ActualHeight+0.5) { throw "Page exceeded available height at $size" }
    $oldWidth=$content.Width
    Update-SystemPageFit -Ui @{ SystemPageHost=$hostPanel; SystemPageContent=$content }
    if ($content.Width -ne $oldWidth) { throw 'Repeated fit was unstable.' }
}
'PASS: full-width layout fits height and stays stable across desktop, laptop, scaled, and compact viewports.'
