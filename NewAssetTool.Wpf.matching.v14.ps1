[CmdletBinding()]
param(
    [string]$XamlPath = $(if ($PSScriptRoot) { Join-Path $PSScriptRoot 'NewAssetTool.matching.v14.xaml' } else { 'NewAssetTool.matching.v3.xaml' })
)

Set-StrictMode -Version 3
$ErrorActionPreference = 'Stop'

Add-Type -AssemblyName PresentationCore
Add-Type -AssemblyName PresentationFramework
Add-Type -AssemblyName WindowsBase

function Show-StartupError {
    param(
        [Parameter(Mandatory=$true)]
        [System.Exception]$Exception,
        [string]$ScriptPath
    )
    $baseFolder = if ($ScriptPath) { Split-Path -Parent $ScriptPath } else { [Environment]::GetFolderPath('Desktop') }
    $logPath = Join-Path $baseFolder 'NewAssetTool.startup-error.log'
    $message = @"
NewAssetTool.Wpf.matching.v14.ps1 failed to start.

Error:
$($Exception.ToString())

Log file:
$logPath
"@
    try { $message | Set-Content -LiteralPath $logPath -Encoding UTF8 } catch {}
    try { [System.Windows.MessageBox]::Show($message, 'New Inventory Tool startup error') | Out-Null } catch { Write-Host $message }
}

try {
    if ([Threading.Thread]::CurrentThread.GetApartmentState() -ne 'STA') {
        throw "This WPF script must be run in STA mode."
    }

    function ConvertFrom-XamlFile {
        param([Parameter(Mandatory=$true)][string]$Path)
        if (-not (Test-Path -LiteralPath $Path)) { throw "XAML file not found: $Path" }
        $rawXaml = Get-Content -LiteralPath $Path -Raw
        $xml = [xml]$rawXaml
        $reader = New-Object System.Xml.XmlNodeReader $xml
        return [Windows.Markup.XamlReader]::Load($reader)
    }

    function Get-NamedControls {
        param(
            [Parameter(Mandatory=$true)][System.Windows.Window]$Window,
            [Parameter(Mandatory=$true)][string[]]$Names
        )
        $controls = @{}
        foreach ($name in $Names) {
            $control = $Window.FindName($name)
            if ($null -eq $control) { throw "Required control '$name' was not found in the XAML." }
            $controls[$name] = $control
        }
        return $controls
    }

    function New-Brush {
        param([Parameter(Mandatory=$true)][string]$Hex)
        $converter = New-Object System.Windows.Media.BrushConverter
        return $converter.ConvertFromString($Hex)
    }

    function Set-BadgeText {
        param([System.Windows.Controls.Border]$Border,[string]$Text)
        if ($Border.Child -is [System.Windows.Controls.TextBlock]) { $Border.Child.Text = $Text }
    }

    function Set-BadgeStyle {
        param([System.Windows.Controls.Border]$Border,[string]$BackgroundHex,[string]$ForegroundHex)
        $Border.Background = New-Brush $BackgroundHex
        if ($Border.Child -is [System.Windows.Controls.TextBlock]) { $Border.Child.Foreground = New-Brush $ForegroundHex }
    }

    function Set-StatusMessage {
        param([hashtable]$Ui,[ValidateSet('Found','PingComplete','Warning')][string]$Mode,[string]$CustomText)
        switch ($Mode) {
            'Found' {
                Set-BadgeText -Border $Ui.StatusMessageBadge -Text $(if ($CustomText) { $CustomText } else { 'Found Computer / Computer' })
                Set-BadgeStyle -Border $Ui.StatusMessageBadge -BackgroundHex '#FCE3E5' -ForegroundHex '#BE123C'
            }
            'PingComplete' {
                Set-BadgeText -Border $Ui.StatusMessageBadge -Text $(if ($CustomText) { $CustomText } else { 'Ping complete' })
                Set-BadgeStyle -Border $Ui.StatusMessageBadge -BackgroundHex '#DDF7E5' -ForegroundHex '#15803D'
            }
            'Warning' {
                Set-BadgeText -Border $Ui.StatusMessageBadge -Text $(if ($CustomText) { $CustomText } else { 'No matching device found' })
                Set-BadgeStyle -Border $Ui.StatusMessageBadge -BackgroundHex '#FEE7C3' -ForegroundHex '#B45309'
            }
        }
        $script:AppState.LastStatusMode = $Mode
    }

    function Get-OutputFolder { param([string]$ResolvedXamlPath) Join-Path (Split-Path -Parent $ResolvedXamlPath) 'Output' }
    function Get-RoundingEventsPath { param([string]$ResolvedXamlPath) Join-Path (Get-OutputFolder -ResolvedXamlPath $ResolvedXamlPath) 'RoundingEvents.csv' }

    function Ensure-OutputFolder {
        param([string]$ResolvedXamlPath)
        $folder = Get-OutputFolder -ResolvedXamlPath $ResolvedXamlPath
        if (-not (Test-Path -LiteralPath $folder)) { $null = New-Item -ItemType Directory -Path $folder -Force }
    }

    function Add-CsvRow {
        param([string]$Path,[psobject]$Row)
        if (Test-Path -LiteralPath $Path) { $Row | Export-Csv -LiteralPath $Path -NoTypeInformation -Append }
        else { $Row | Export-Csv -LiteralPath $Path -NoTypeInformation }
    }

    function Set-DisplayText {
        param([hashtable]$Ui,[string]$BaseName,[string]$Value)
        $display = $Ui["${BaseName}Display"]
        if ($display) { $display.Text = $Value }
        $textbox = $Ui["${BaseName}TextBox"]
        if ($textbox) { $textbox.Text = $Value }
    }

    function New-SampleData {
        $device = [pscustomobject]@{
            SearchKeys   = @('AO400568', 'HSS-8093577', 'C24102M031')
            DetectedType = 'Tangent'
            Name         = 'AO400568'
            AssetTag     = 'HSS-8093577'
            Serial       = 'C24102M031'
            Parent       = '(n/a)'
            RITM         = 'TRP - 26 May 2025'
            RetireDate   = '31 May 2028'
            LastRounded  = '12 March 2026 - 39 days ago'
            City         = 'Duncan'
            Location     = 'VIHA-CDH-Cowichan District Hospital'
            Building     = 'Main Building'
            Floor        = '1'
            Room         = '1068 (PACU)'
            Department   = 'Medical Device Reprocessing Department (MDRD)'
        }

        $associated = @(
            [pscustomobject]@{ Role='Parent'; Type='Tangent'; Name='AO400568'; AssetTag='HSS-8093577'; Serial='C24102M031'; RITM='TRP - 26 May 2025'; Retire='31 May 2028' },
            [pscustomobject]@{ Role='Child'; Type='Cart'; Name='AO400568-CRT'; AssetTag='CO09167'; Serial='1896875-0016'; RITM='-'; Retire='-' }
        )

        $nearby = @(
            [pscustomobject]@{ HostName='LD065898'; IPAddress='';             Subnet='';       AssetTag='HSS-8077199'; Location='VIHA-DNDR-Duncan Norcr...'; Building='Main Building'; Floor='1'; Room='101 (#6 Charge Cabinet)'; Department='CHS - Community Health Se...'; MaintenanceType='General Rounding'; LastRounded='20 April 2026'; DaysAgo='0';   Status='Inaccessible - Asset not found' },
            [pscustomobject]@{ HostName='LD065911'; IPAddress='10.64.45.232'; Subnet='VPN';    AssetTag='HSS-8077204'; Location='VIHA-DNDR-Duncan Norcr...'; Building='Main Building'; Floor='1'; Room='101 (#7 Charge Cabinet)'; Department='CHS - Community Health Se...'; MaintenanceType='General Rounding'; LastRounded='20 April 2026'; DaysAgo='0';   Status='Inaccessible - Laptop is not available' },
            [pscustomobject]@{ HostName='LD062047'; IPAddress='10.64.47.15';  Subnet='VPN';    AssetTag='HSS-1037495'; Location='VIHA-DNDR-Duncan Norcr...'; Building='Main Building'; Floor='1'; Room='101 (#8 Charge Cabinet)'; Department='CHS (Reception)'; MaintenanceType='General Rounding'; LastRounded='20 April 2026'; DaysAgo='0'; Status='Inaccessible - Laptop is not available' },
            [pscustomobject]@{ HostName='PC077708'; IPAddress='10.209.233.167';Subnet='Unknown';AssetTag='HSS-1037501'; Location='VIHA-DNDR-Duncan Norcr...'; Building='Main Building'; Floor='1'; Room='102'; Department='CHS - Community Health Se...'; MaintenanceType='General Rounding'; LastRounded='06 March 2026'; DaysAgo='45'; Status='-' },
            [pscustomobject]@{ HostName='LD072236'; IPAddress='10.209.233.47'; Subnet='Unknown';AssetTag='HSS-1037488'; Location='VIHA-DNDR-Duncan Norcr...'; Building='Main Building'; Floor='1'; Room='104 (Chart Room)'; Department='Charting'; MaintenanceType='General Rounding'; LastRounded='20 April 2026'; DaysAgo='0'; Status='Complete' }
        )

        return [pscustomobject]@{ Device=$device; Associated=$associated; Nearby=$nearby }
    }

    function Set-WindowDataBindings {
        param([hashtable]$Ui,[pscustomobject]$SampleData)
        $device = $SampleData.Device
        $Ui.SearchTextBox.Text = $device.Name
        $Ui.SelectedDeviceText.Text = $device.Name
        Set-DisplayText -Ui $Ui -BaseName 'DetectedType' -Value $device.DetectedType
        Set-DisplayText -Ui $Ui -BaseName 'HostName' -Value $device.Name
        Set-DisplayText -Ui $Ui -BaseName 'AssetTag' -Value $device.AssetTag
        Set-DisplayText -Ui $Ui -BaseName 'Serial' -Value $device.Serial
        Set-DisplayText -Ui $Ui -BaseName 'Parent' -Value $device.Parent
        Set-DisplayText -Ui $Ui -BaseName 'Ritm' -Value $device.RITM
        Set-DisplayText -Ui $Ui -BaseName 'Retire' -Value $device.RetireDate
        $Ui.LastRoundedText.Text = $device.LastRounded
        $Ui.CityTextBox.Text = $device.City
        $Ui.LocationTextBox.Text = $device.Location
        $Ui.BuildingTextBox.Text = $device.Building
        $Ui.FloorTextBox.Text = $device.Floor
        $Ui.RoomTextBox.Text = $device.Room
        $Ui.DepartmentTextBox.Text = $device.Department
        $Ui.LastQueryBadgeText.Text = "Queried $(Get-Date -Format 'HH:mm')"
        $Ui.NearbyScopeSummaryText.Text = "Nearby scopes (Location): 1 - Showing $($SampleData.Nearby.Count)"
        $Ui.AssociatedDevicesDataGrid.ItemsSource = $SampleData.Associated
        $Ui.NearbyDataGrid.ItemsSource = $SampleData.Nearby
    }

    function Find-SampleDevice {
        param([string]$SearchTerm,[pscustomobject]$SampleData)
        $term = $SearchTerm.Trim()
        if ([string]::IsNullOrWhiteSpace($term)) { return $null }
        foreach ($key in $SampleData.Device.SearchKeys) {
            if ($key -like "*$term*" -or $term -like "*$key*") { return $SampleData.Device }
        }
        return $null
    }

    function Get-RoundingMinutes {
        param([hashtable]$Ui)
        $minutes = 0
        if (-not [int]::TryParse($Ui.RoundingTimeTextBox.Text, [ref]$minutes)) { return 0 }
        return [Math]::Max(0, $minutes)
    }

    function Set-RoundingMinutes {
        param([hashtable]$Ui,[int]$Minutes)
        $Ui.RoundingTimeTextBox.Text = [Math]::Max(0, $Minutes).ToString()
    }

    function Save-RoundingEvent {
        param([hashtable]$Ui,[pscustomobject]$CurrentDevice,[string]$ResolvedXamlPath)
        Ensure-OutputFolder -ResolvedXamlPath $ResolvedXamlPath
        $csvPath = Get-RoundingEventsPath -ResolvedXamlPath $ResolvedXamlPath
        $row = [pscustomobject]@{
            AssetTag=$CurrentDevice.AssetTag; Name=$CurrentDevice.Name; Serial=$CurrentDevice.Serial
            City=$CurrentDevice.City; Location=$CurrentDevice.Location; Building=$CurrentDevice.Building
            Floor=$CurrentDevice.Floor; Room=$CurrentDevice.Room; CheckStatus=$Ui.CheckStatusComboBox.Text
            RoundingMinutes=(Get-RoundingMinutes -Ui $Ui); CableMgmtOK=[bool]$Ui.ValidateCableCheckBox.IsChecked
            LabelOK=[bool]$Ui.LabelMonitorCheckBox.IsChecked; CartOK=[bool]$Ui.PhysicalCartCheckBox.IsChecked
            PeripheralsOK=[bool]$Ui.ValidatePeripheralsCheckBox.IsChecked; Comments=$Ui.CommentsTextBox.Text
            SavedAt=(Get-Date).ToString('s'); SavedFrom='System'
        }
        Add-CsvRow -Path $csvPath -Row $row
        [System.Windows.MessageBox]::Show("Saved event to:`n$csvPath", 'Save Event') | Out-Null
    }

    function Save-NearbyEvents {
        param([hashtable]$Ui,[string]$ResolvedXamlPath)
        Ensure-OutputFolder -ResolvedXamlPath $ResolvedXamlPath
        $csvPath = Get-RoundingEventsPath -ResolvedXamlPath $ResolvedXamlPath
        $selectedRows = @($Ui.NearbyDataGrid.SelectedItems)
        if ($selectedRows.Count -eq 0) {
            [System.Windows.MessageBox]::Show('Select one or more Nearby rows before saving.', 'Nearby Save') | Out-Null
            return
        }
        foreach ($item in $selectedRows) {
            $row = [pscustomobject]@{
                AssetTag=$item.AssetTag; Name=$item.HostName; Serial=''; City='Duncan'; Location=$item.Location; Building=$item.Building; Floor=$item.Floor; Room=$item.Room
                CheckStatus=$(if ([string]::IsNullOrWhiteSpace($item.Status) -or $item.Status -eq '-') { 'Complete' } else { $item.Status })
                RoundingMinutes=3; CableMgmtOK=$true; LabelOK=$true; CartOK=$true; PeripheralsOK=$true
                Comments='Saved from Nearby tab'; SavedAt=(Get-Date).ToString('s'); SavedFrom='Nearby'
            }
            Add-CsvRow -Path $csvPath -Row $row
        }
        [System.Windows.MessageBox]::Show("Saved $($selectedRows.Count) nearby row(s) to:`n$csvPath", 'Nearby Save') | Out-Null
    }

    function Toggle-LocationEditMode {
        param([hashtable]$Ui,[bool]$IsEditing)
        $readOnlyControls = @($Ui.CityTextBox,$Ui.LocationTextBox,$Ui.BuildingTextBox,$Ui.FloorTextBox,$Ui.RoomTextBox,$Ui.DepartmentTextBox)
        $comboControls = @($Ui.CityComboBox,$Ui.LocationComboBox,$Ui.BuildingComboBox,$Ui.FloorComboBox,$Ui.RoomComboBox,$Ui.DepartmentComboBox)
        foreach ($control in $readOnlyControls) { $control.Visibility = if ($IsEditing) { 'Collapsed' } else { 'Visible' } }
        foreach ($control in $comboControls) { $control.Visibility = if ($IsEditing) { 'Visible' } else { 'Collapsed' } }
        $Ui.CancelEditLocationButton.Visibility = if ($IsEditing) { 'Visible' } else { 'Collapsed' }
        $Ui.EditLocationButton.Content = if ($IsEditing) { 'Save' } else { 'Edit Location' }
    }

    function Save-LocationValues {
        param([hashtable]$Ui)

        $Ui.CityTextBox.Text = $Ui.CityComboBox.Text
        $Ui.LocationTextBox.Text = $Ui.LocationComboBox.Text
        $Ui.BuildingTextBox.Text = $Ui.BuildingComboBox.Text
        $Ui.FloorTextBox.Text = $Ui.FloorComboBox.Text
        $Ui.RoomTextBox.Text = $Ui.RoomComboBox.Text
        $Ui.DepartmentTextBox.Text = $Ui.DepartmentComboBox.Text
    }

    $resolvedXamlPath = (Resolve-Path -LiteralPath $XamlPath).Path
    $window = ConvertFrom-XamlFile -Path $resolvedXamlPath

    $ui = Get-NamedControls -Window $window -Names @(
        'SearchTextBox','QueryButton','PingButton','LiveDetailsButton','MonitorLabelButton',
        'MainTabControl','SystemTab','NearbyTab','SelectedDeviceText','DeviceOnlineText','LastQueryBadgeText',
        'DetectedTypeDisplay','HostNameDisplay','AssetTagDisplay','SerialDisplay','ParentDisplay','RitmDisplay','RetireDisplay',
        'DetectedTypeTextBox','HostNameTextBox','AssetTagTextBox','SerialNumberTextBox','ParentTextBox','RitmTextBox','RetireDateTextBox','LastRoundedText',
        'AssociatedDevicesDataGrid','FixNameButton',
        'AddPeripheralButton','RemovePeripheralButton','ValidateAssociatedButton',
        'CityTextBox','LocationTextBox','BuildingTextBox','FloorTextBox','RoomTextBox','DepartmentTextBox',
        'CityComboBox','LocationComboBox','BuildingComboBox','FloorComboBox','RoomComboBox','DepartmentComboBox',
        'EditLocationButton','CancelEditLocationButton',
        'MaintenanceTypeComboBox','CheckStatusComboBox','RoundingTimeTextBox',
        'RoundingTimeUpButton','RoundingTimeDownButton',
        'ValidateCableCheckBox','LabelMonitorCheckBox','ValidatePeripheralsCheckBox',
        'CablingNeededCheckBox','PhysicalCartCheckBox','AddDeviceToTrackerCheckBox',
        'CheckCompleteButton','SaveEventButton','ManualRoundButton','CommentsTextBox',
        'NearbyScopeSummaryText','RebuildNearbyButton','PingAllButton','ClearNearbyButton',
        'NearbyCollapseButton','NearbyExpandButton','NearbyDataGrid','NearbySaveButton',
        'ShowAllNearbyCheckBox','TodaysRoundedCheckBox','ExcludedCheckBox','RecentlyRoundedCheckBox','CriticalClinicalCheckBox',
        'DataPathText','OutputPathText','DaysPerWeekBadge','TodayBadge','ThisWeekBadge','RemainingPerDayBadge','StatusMessageBadge'
    )

    $script:AppState = [pscustomobject]@{ LastStatusMode='Found'; SampleData=(New-SampleData); CurrentDevice=$null }
    $script:AppState.CurrentDevice = $script:AppState.SampleData.Device

    Set-WindowDataBindings -Ui $ui -SampleData $script:AppState.SampleData
    Set-StatusMessage -Ui $ui -Mode 'Found'
    Toggle-LocationEditMode -Ui $ui -IsEditing:$false

    $ui.DataPathText.Text = "Data: $(Join-Path (Split-Path -Parent $resolvedXamlPath) 'Data')"
    $ui.OutputPathText.Text = "Output: $(Get-OutputFolder -ResolvedXamlPath $resolvedXamlPath)"

    foreach ($pair in @(
        @{ Combo=$ui.CityComboBox; Value=$script:AppState.CurrentDevice.City },
        @{ Combo=$ui.LocationComboBox; Value=$script:AppState.CurrentDevice.Location },
        @{ Combo=$ui.BuildingComboBox; Value=$script:AppState.CurrentDevice.Building },
        @{ Combo=$ui.FloorComboBox; Value=$script:AppState.CurrentDevice.Floor },
        @{ Combo=$ui.RoomComboBox; Value=$script:AppState.CurrentDevice.Room },
        @{ Combo=$ui.DepartmentComboBox; Value=$script:AppState.CurrentDevice.Department }
    )) {
        $pair.Combo.Items.Clear()
        [void]$pair.Combo.Items.Add($pair.Value)
        [void]$pair.Combo.Items.Add('2nd Choice')
        $pair.Combo.SelectedIndex = 0
    }

    $ui.QueryButton.Add_Click({
        $match = Find-SampleDevice -SearchTerm $ui.SearchTextBox.Text -SampleData $script:AppState.SampleData
        if ($null -eq $match) {
            Set-StatusMessage -Ui $ui -Mode 'Warning' -CustomText 'No matching device found'
            return
        }
        $script:AppState.CurrentDevice = $match
        Set-WindowDataBindings -Ui $ui -SampleData $script:AppState.SampleData
        Set-StatusMessage -Ui $ui -Mode 'Found'
    })

    $ui.SearchTextBox.Add_KeyDown({
        if ($_.Key -eq 'Return') {
            $ui.QueryButton.RaiseEvent((New-Object System.Windows.RoutedEventArgs([System.Windows.Controls.Button]::ClickEvent)))
        }
    })

    $ui.PingButton.Add_Click({ [System.Windows.MessageBox]::Show('Ping button clicked.', 'Ping') | Out-Null })
    $ui.LiveDetailsButton.Add_Click({ [System.Windows.MessageBox]::Show('Live Details button clicked.', 'Live Details') | Out-Null })
    $ui.MonitorLabelButton.Add_Click({ [System.Windows.MessageBox]::Show('Monitor Label button clicked.', 'Monitor Label') | Out-Null })
    $ui.FixNameButton.Add_Click({ [System.Windows.MessageBox]::Show('Fix Name button clicked.', 'Fix Name') | Out-Null })
    $ui.EditLocationButton.Add_Click({
        if ($ui.EditLocationButton.Content -eq 'Save') {
            Save-LocationValues -Ui $ui
            Toggle-LocationEditMode -Ui $ui -IsEditing:$false
        }
        else {
            Toggle-LocationEditMode -Ui $ui -IsEditing:$true
        }
    })
    $ui.CancelEditLocationButton.Add_Click({ Toggle-LocationEditMode -Ui $ui -IsEditing:$false })
    $ui.RoundingTimeUpButton.Add_Click({ Set-RoundingMinutes -Ui $ui -Minutes ((Get-RoundingMinutes -Ui $ui) + 1) })
    $ui.RoundingTimeDownButton.Add_Click({ Set-RoundingMinutes -Ui $ui -Minutes ([Math]::Max(0, (Get-RoundingMinutes -Ui $ui) - 1)) })
    $ui.CheckCompleteButton.Add_Click({ $ui.CheckStatusComboBox.Text = 'Complete' })
    $ui.SaveEventButton.Add_Click({ Save-RoundingEvent -Ui $ui -CurrentDevice $script:AppState.CurrentDevice -ResolvedXamlPath $resolvedXamlPath })
    $ui.ManualRoundButton.Add_Click({ [System.Windows.MessageBox]::Show('Manual Round button clicked.', 'Manual Round') | Out-Null })
    $ui.RebuildNearbyButton.Add_Click({ $ui.NearbyScopeSummaryText.Text = "Nearby scopes (Location): 1 - Showing $($script:AppState.SampleData.Nearby.Count)" })
    $ui.ClearNearbyButton.Add_Click({ $ui.NearbyDataGrid.UnselectAll() })
    $ui.PingAllButton.Add_Click({ Set-StatusMessage -Ui $ui -Mode 'PingComplete' })
    $ui.NearbySaveButton.Add_Click({ Save-NearbyEvents -Ui $ui -ResolvedXamlPath $resolvedXamlPath })

    $ui.MainTabControl.Add_SelectionChanged({
        if ($window.IsLoaded) {
            if ($script:AppState.LastStatusMode -eq 'PingComplete') { Set-StatusMessage -Ui $ui -Mode 'PingComplete' }
            else { Set-StatusMessage -Ui $ui -Mode 'Found' }
        }
    })

    $ui.ValidatePeripheralsCheckBox.IsChecked = $true
    $ui.LabelMonitorCheckBox.IsChecked = $false
    $ui.ValidateCableCheckBox.IsChecked = $false
    $ui.CablingNeededCheckBox.IsChecked = $false
    $ui.PhysicalCartCheckBox.IsChecked = $false
    $ui.AddDeviceToTrackerCheckBox.IsChecked = $false

    [void]$window.ShowDialog()
}
catch {
    Show-StartupError -Exception $_.Exception -ScriptPath $PSCommandPath
    throw
}
