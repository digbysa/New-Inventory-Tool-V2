# Run in an isolated PowerShell process. A fake dispatcher exercises the real
# runspace worker without requiring Windows or a WPF desktop.
$ErrorActionPreference = 'Stop'
Set-StrictMode -Version 3
Add-Type @'
namespace System.Windows.Threading {
    public class DispatcherTimer {
        public System.TimeSpan Interval { get; set; }
        public object Tag { get; set; }
        public event System.EventHandler Tick;
        public void Start() { }
        public void Stop() { }
        public void Fire() { Tick?.Invoke(this, System.EventArgs.Empty); }
    }
}
'@
$tokens = $null; $parseErrors = $null
$path = Join-Path $PSScriptRoot '../NewAssetTool.Wpf.matching.v14.ps1'
$ast = [System.Management.Automation.Language.Parser]::ParseFile($path, [ref]$tokens, [ref]$parseErrors)
if ($parseErrors.Count) { throw ($parseErrors | Out-String) }
$functionAst = $ast.Find({ param($node) $node -is [System.Management.Automation.Language.FunctionDefinitionAst] -and $node.Name -eq 'Invoke-CurrentDevicePing' }, $true)
Invoke-Expression $functionAst.Extent.Text
function Resolve-CurrentPingTarget { 'test-host' }
function Invoke-DevicePing { [pscustomobject]@{ Success=$true } }
function Test-RemoteConnectivity { [pscustomobject]@{ HostName='test-host'; IsOnline=$true; LatencyMs=12; IpAddress='10.0.0.1'; Subnet='Test' } }
function Get-RemoteRestartStatus { Start-Sleep -Milliseconds 1500; [pscustomobject]@{ Uptime='1d 2h 3m'; PendingReboot='No' } }
# Imported dependencies are unused by the simulated network helpers.
function Invoke-PingOnce {}
function Get-FieldValue {}
function Get-IPv4AddressFromPingReply {}
function Resolve-HostIPv4Address {}
function Resolve-SubnetName {}
function ConvertTo-IPv4Bytes {}
$script:updates = [System.Collections.ArrayList]::new()
$script:continuousCount = 0
function Set-OnlineStatusUi {
    param($Ui,$IsOnline,$LatencyMs,$IpAddress,$Subnet,$CheckedHost,$Uptime,$PendingReboot)
    [void]$script:updates.Add($Uptime)
}
function Start-ContinuousPingWindow { param($Target) $script:continuousCount++ }
function Set-StatusMessage { param($Ui,$Mode,$CustomText) }
$script:SystemPingTimers = [System.Collections.ArrayList]::new()
$script:AppState = [pscustomobject]@{ CurrentQueryToken='first' }
$script:CurrentSystemPingRequest = ''
$watch = [Diagnostics.Stopwatch]::StartNew()
Invoke-CurrentDevicePing -Ui @{} -Inventory $null -DataRoot '' -StartContinuous
if ($watch.ElapsedMilliseconds -ge 1200) { throw 'UI thread waited for the remote restart query.' }
$deadline = [datetime]::UtcNow.AddSeconds(8)
while ($script:SystemPingTimers.Count -gt 0 -and [datetime]::UtcNow -lt $deadline) {
    foreach ($timer in @($script:SystemPingTimers)) { $timer.Fire() }
    Start-Sleep -Milliseconds 25
}
if ($script:SystemPingTimers.Count) { throw 'Worker was not cleaned up.' }
if ($script:updates -notcontains 'Checking...') { throw 'Connectivity was not published before restart data.' }
if ($script:updates[-1] -ne '1d 2h 3m') { throw 'Restart result was not published.' }
if ($script:continuousCount -ne 1) { throw 'Explicit ping should open exactly one continuous window.' }
$script:updates.Clear()
Invoke-CurrentDevicePing -Ui @{} -Inventory $null -DataRoot ''
$script:AppState.CurrentQueryToken = 'second'
$deadline = [datetime]::UtcNow.AddSeconds(8)
while ($script:SystemPingTimers.Count -gt 0 -and [datetime]::UtcNow -lt $deadline) {
    foreach ($timer in @($script:SystemPingTimers)) { $timer.Fire() }
    Start-Sleep -Milliseconds 25
}
if ($script:updates.Count) { throw 'An old query updated the current device.' }
if ($script:SystemPingTimers.Count) { throw 'Stale worker was not cleaned up.' }
Write-Output 'PASS: script parsing, nonblocking ping, early connectivity, restart result, explicit continuous ping, stale query protection, worker cleanup.'
# A replacement request cancels the previous one; only the newest may publish.
$script:updates.Clear()
Invoke-CurrentDevicePing -Ui @{} -Inventory $null -DataRoot ''
Invoke-CurrentDevicePing -Ui @{} -Inventory $null -DataRoot ''
$deadline = [datetime]::UtcNow.AddSeconds(8)
while ($script:SystemPingTimers.Count -gt 0 -and [datetime]::UtcNow -lt $deadline) {
    foreach ($timer in @($script:SystemPingTimers)) { $timer.Fire() }
    Start-Sleep -Milliseconds 25
}
if ($script:SystemPingTimers.Count) { throw 'Cancelled worker was not cleaned up.' }
if (@($script:updates | Where-Object { $_ -eq '1d 2h 3m' }).Count -ne 1) { throw 'Replacement request published incorrect results.' }
# Verify field lookup semantics using the real implementation.
$fieldAst = $ast.Find({ param($node) $node -is [System.Management.Automation.Language.FunctionDefinitionAst] -and $node.Name -eq 'Get-FieldValue' }, $true)
Invoke-Expression $fieldAst.Extent.Text
$row = [pscustomobject]@{ Empty=$null; Blank=' '; HostName='PC001'; Zero=0 }
if ((Get-FieldValue -Row $row -Names @('missing','Empty','Blank','hostname')) -ne 'PC001') { throw 'Field fallback/case handling changed.' }
if ((Get-FieldValue -Row $row -Names @('Zero')) -ne '0') { throw 'Zero field was discarded.' }
if ((Get-FieldValue -Row $row -Names @('missing')) -ne '') { throw 'Missing field did not return empty text.' }
Write-Output 'PASS: replacement request cancellation and field lookup semantics.'
