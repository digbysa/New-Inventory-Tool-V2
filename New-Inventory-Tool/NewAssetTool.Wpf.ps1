Add-Type -AssemblyName PresentationCore,PresentationFramework,WindowsBase,System.Xaml
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName WindowsFormsIntegration

function Get-LoaderScriptDir {
  try {
    if ($PSScriptRoot -and $PSScriptRoot -ne '') { return $PSScriptRoot }
  } catch {}
  try {
    if ($PSCommandPath -and $PSCommandPath -ne '') { return (Split-Path -Parent $PSCommandPath) }
  } catch {}
  try {
    if ($MyInvocation -and $MyInvocation.MyCommand -and $MyInvocation.MyCommand.Path) {
      return (Split-Path -Parent $MyInvocation.MyCommand.Path)
    }
  } catch {}
  return (Get-Location).Path
}

$scriptDir = Get-LoaderScriptDir
$ps1Path   = Join-Path $scriptDir 'NewAssetTool.ps1'
$xamlPath  = Join-Path $scriptDir 'NewAssetTool.xaml'

if (-not (Test-Path $ps1Path)) {
  throw "Could not find NewAssetTool.ps1 at '$ps1Path'."
}
if (-not (Test-Path $xamlPath)) {
  throw "Could not find NewAssetTool.xaml at '$xamlPath'."
}

$previousSuppress = $null
$hadPrevious = $false
try {
  if (Test-Path Variable:\global:NewAssetToolSuppressShow) {
    $previousSuppress = $global:NewAssetToolSuppressShow
    $hadPrevious = $true
  }
} catch {}
$global:NewAssetToolSuppressShow = $true

try {
  $scriptOutput = . $ps1Path
} catch {
  if ($hadPrevious) {
    $global:NewAssetToolSuppressShow = $previousSuppress
  } else {
    Remove-Variable -Scope Global -Name NewAssetToolSuppressShow -ErrorAction SilentlyContinue
  }
  throw
}

if ($scriptOutput -is [System.Windows.Forms.Form]) {
  $form = $scriptOutput
} else {
  $form = ($scriptOutput | Where-Object { $_ -is [System.Windows.Forms.Form] } | Select-Object -First 1)
}

if (-not $form -and $script:NewAssetToolMainForm -is [System.Windows.Forms.Form]) {
  $form = $script:NewAssetToolMainForm
}

if (-not $form -or -not ($form -is [System.Windows.Forms.Form])) {
  throw "NewAssetTool.ps1 did not expose a main form instance."
}

$form.TopLevel = $false
$form.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::None
$form.Dock = [System.Windows.Forms.DockStyle]::Fill

if (-not (Test-Path $xamlPath)) {
  throw "XAML window definition missing at '$xamlPath'."
}

$reader = [System.Xml.XmlReader]::Create($xamlPath)
try {
  $window = [System.Windows.Markup.XamlReader]::Load($reader)
} finally {
  $reader.Close()
}

try {
  $iconPath = Join-Path $scriptDir 'icon.ico'
  if (Test-Path $iconPath) {
    $iconImage = New-Object System.Windows.Media.Imaging.BitmapImage
    $iconImage.BeginInit()
    $iconImage.UriSource = New-Object System.Uri($iconPath)
    $iconImage.CacheOption = [System.Windows.Media.Imaging.BitmapCacheOption]::OnLoad
    $iconImage.EndInit()
    $window.Icon = $iconImage
  }
} catch {}

$selectedSiteName = $null
try { $selectedSiteName = $script:SelectedSiteName } catch {}
if ([string]::IsNullOrWhiteSpace($selectedSiteName) -and $form -and $form.Text) {
  $selectedSiteName = ($form.Text -replace '^New Inventory Tool\s*-\s*', '').Trim()
}
if ($window) {
  $window.DataContext = [pscustomobject]@{
    SelectedSiteName = $selectedSiteName
  }
}

$windowsFormsHost = $window.FindName('WinFormsHost')
if (-not $windowsFormsHost) {
  throw "Could not locate the WindowsFormsHost named 'WinFormsHost' in XAML."
}
$windowsFormsHost.Child = $form
$form.Visible = $true

$rootGrid = $window.FindName('RootGrid')
$rootScaleTransform = $window.FindName('RootScaleTransform')
$applyWpfScale = {
  param([string]$Source = 'unspecified')

  if (-not $rootScaleTransform) {
    Write-Verbose "[DPI][WPF] Root scale transform not located; skipping scale." -Verbose
    return
  }

  $scale = 1.0
  try {
    if (Get-Command Get-NewAssetToolChromeScale -ErrorAction SilentlyContinue) {
      $scale = Get-NewAssetToolChromeScale
    }
  } catch {}

  if (-not $global:NewAssetToolPerMonitorDpiContextEnabled -and [Math]::Abs($scale - 1.0) -lt 0.0001) {
    return
  }

  $contextDescription = 'unknown'
  try {
    if (Get-Command Get-NewAssetToolDpiContextDescription -ErrorAction SilentlyContinue) {
      $contextDescription = Get-NewAssetToolDpiContextDescription
    }
  } catch {}

  try {
    $rootScaleTransform.ScaleX = $scale
    $rootScaleTransform.ScaleY = $scale
    Write-Verbose (
      "[DPI][WPF] Applied root layout scale {0:n3} ({1}) context={2}" -f $scale, $Source, $contextDescription
    ) -Verbose
  } catch {
    Write-Verbose "[DPI][WPF] Failed to apply layout scale ({0}): $($_.Exception.Message)" -Verbose
  }
}

Set-Variable -Scope Global -Name NewAssetToolApplyWpfScale -Value $applyWpfScale

if ($window) {
  if ($global:NewAssetToolPerMonitorDpiContextEnabled) {
    try {
      if (Get-Command Set-NewAssetToolMonitorScale -ErrorAction SilentlyContinue) {
        $dpiInfo = [System.Windows.Media.VisualTreeHelper]::GetDpi($window)
        $initialScale = $null
        try { $initialScale = [double]$dpiInfo.DpiScaleX } catch {}
        if ($null -ne $initialScale -and $initialScale -gt 0) {
          Set-NewAssetToolMonitorScale -Scale $initialScale -Source 'Wpf.Initial' | Out-Null
        }
      }
    } catch {}
  }

  if ($global:NewAssetToolPerMonitorDpiContextEnabled) {
    try {
      $window.Add_DpiChanged({
        param($sender,$eventArgs)

        if (Get-Command Set-NewAssetToolMonitorScale -ErrorAction SilentlyContinue) {
          $scale = $null
          try {
            $dpiScale = $eventArgs.NewDpi.DpiScaleX
            if ($dpiScale -gt 0) { $scale = [double]$dpiScale }
          } catch {}
          if ($null -eq $scale) {
            try {
              $pixelsPerDip = $eventArgs.NewDpi.PixelsPerDip
              if ($pixelsPerDip -gt 0) { $scale = [double]$pixelsPerDip }
            } catch {}
          }
          if ($null -ne $scale -and $scale -gt 0) {
            Set-NewAssetToolMonitorScale -Scale $scale -Source 'Wpf.DpiChanged' | Out-Null
          }
        }

        & $applyWpfScale 'Window.DpiChanged'
      })
    } catch {
      Write-Verbose "[DPI][WPF] Failed to attach DpiChanged handler: $($_.Exception.Message)" -Verbose
    }
  }
  & $applyWpfScale 'Window.Initial'
}

if ($window) {
  if ($window.WindowState -ne [System.Windows.WindowState]::Normal) {
    $window.WindowState = [System.Windows.WindowState]::Normal
  }
  if ($window.SizeToContent -ne [System.Windows.SizeToContent]::Height) {
    $window.SizeToContent = [System.Windows.SizeToContent]::Height
  }
  $sizeLockApplied = $false
  $window.Add_ContentRendered({
    param($sender, $eventArgs)
    if (-not $sizeLockApplied) {
      if ($sender.SizeToContent -ne [System.Windows.SizeToContent]::Manual) {
        $sender.SizeToContent = [System.Windows.SizeToContent]::Manual
      }
      $sizeLockApplied = $true
    }
  })
}

$searchTextBox = $window.FindName('SearchTextBox')
$lookupButton = $window.FindName('LookupButton')
$bannerDeviceNameText = $window.FindName('BannerDeviceNameText')

function Update-BannerDeviceName {
  $bannerText = ''
  try {
    if ($txtParent) {
      $bannerText = ('' + $txtParent.Text).Trim()
    }
  } catch {}
  if ([string]::IsNullOrWhiteSpace($bannerText)) {
    try {
      if ($script:CurrentParent -and $script:CurrentParent.name) {
        $bannerText = ('' + $script:CurrentParent.name).Trim()
      }
    } catch {}
  }
  if ([string]::IsNullOrWhiteSpace($bannerText)) {
    $bannerText = '-'
  }
  try {
    if ($bannerDeviceNameText) {
      $bannerDeviceNameText.Text = $bannerText
    }
  } catch {}
}

if ($searchTextBox) {
  try { Set-ScanSearchControl $searchTextBox } catch {}
  $hasHandleScanTextChanged = $false
  $hasUpdateRoundNowButtonState = $false
  try { $hasHandleScanTextChanged = [bool](Get-Command Handle-ScanTextChanged -ErrorAction SilentlyContinue) } catch {}
  try { $hasUpdateRoundNowButtonState = [bool](Get-Command Update-RoundNowButtonState -ErrorAction SilentlyContinue) } catch {}
  $scanSyncTimer = New-Object System.Windows.Threading.DispatcherTimer
  $scanSyncTimer.Interval = [TimeSpan]::FromMilliseconds(60)
  $scanSyncTimer.Add_Tick({
    $scanSyncTimer.Stop()
    if ($txtScan) {
      if ($txtScan.Text -ne $script:PendingScanText) {
        $txtScan.Text = $script:PendingScanText
      }
    }
  })
  if ($txtScan) {
    if ($searchTextBox.Text -ne $txtScan.Text) { $searchTextBox.Text = $txtScan.Text }
    $txtScan.Add_TextChanged({
      param($sender,$eventArgs)
      $target = $script:NewAssetToolSearchTextBox
      if ($target -and $target.Text -ne $sender.Text) {
        $target.Text = $sender.Text
        try { $target.CaretIndex = $target.Text.Length } catch {}
      }
    })
  }
  $searchTextBox.Add_TextChanged({
    param($sender,$eventArgs)
    $script:PendingScanText = $sender.Text
    $scanSyncTimer.Stop()
    $scanSyncTimer.Start()
    if ($hasHandleScanTextChanged) {
      Handle-ScanTextChanged $sender.Text
    }
    try {
      if ($hasUpdateRoundNowButtonState) {
        Update-RoundNowButtonState
      }
    } catch {}
  })
  $searchTextBox.Add_KeyDown({
    param($sender,$eventArgs)
    if ($eventArgs.Key -in @([System.Windows.Input.Key]::Enter,[System.Windows.Input.Key]::Return)) {
      $eventArgs.Handled = $true
      try { Do-Lookup } catch {}
      try { Update-BannerDeviceName } catch {}
    }
  })
  try { Focus-ScanInput } catch {}
  try {
    if ($hasUpdateRoundNowButtonState) {
      Update-RoundNowButtonState
    }
  } catch {}
}

if ($lookupButton) {
  $lookupButton.Add_Click({
    try { Do-Lookup } catch {}
    try { Update-BannerDeviceName } catch {}
  })
}

if ($txtParent) {
  $txtParent.Add_TextChanged({
    try { Update-BannerDeviceName } catch {}
  })
}

try { Update-BannerDeviceName } catch {}

$app = [System.Windows.Application]::Current
if (-not $app) { $app = New-Object System.Windows.Application }
$app.ShutdownMode = [System.Windows.ShutdownMode]::OnMainWindowClose
$app.MainWindow = $window

$script:__newAssetToolCleanupRan = $false
$cleanupAction = {
  if (-not $script:__newAssetToolCleanupRan) {
    $script:__newAssetToolCleanupRan = $true
    try { if ($form) { $form.Dispose() } } catch {}
    try {
      if ($hadPrevious) {
        $global:NewAssetToolSuppressShow = $previousSuppress
      } else {
        Remove-Variable -Scope Global -Name NewAssetToolSuppressShow -ErrorAction SilentlyContinue
      }
    } catch {}
    try { Remove-Variable -Scope Global -Name NewAssetToolApplyWpfScale -ErrorAction SilentlyContinue } catch {}
  }
}

$window.Add_Closed($cleanupAction)

try {
  [void]$app.Run($window)
} finally {
  & $cleanupAction
}
