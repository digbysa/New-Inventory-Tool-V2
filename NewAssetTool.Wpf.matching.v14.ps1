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

    function Get-DataFolder { param([string]$ResolvedXamlPath) Join-Path (Split-Path -Parent $ResolvedXamlPath) 'Data' }
    function Get-OutputFolder { param([string]$ResolvedXamlPath) Join-Path (Split-Path -Parent $ResolvedXamlPath) 'Output' }
    function Get-RoundingEventsPath { param([string]$ResolvedXamlPath) Join-Path (Get-OutputFolder -ResolvedXamlPath $ResolvedXamlPath) 'RoundingEvents.csv' }

    function Ensure-OutputFolder {
        param([string]$ResolvedXamlPath)
        $folder = Get-OutputFolder -ResolvedXamlPath $ResolvedXamlPath
        if (-not (Test-Path -LiteralPath $folder)) { $null = New-Item -ItemType Directory -Path $folder -Force }
    }

    function Add-CsvRow {
        param([string]$Path,[psobject]$Row)
        if (Test-Path -LiteralPath $Path) { $Row | Export-Csv -LiteralPath $Path -NoTypeInformation -Append -Encoding UTF8 }
        else { $Row | Export-Csv -LiteralPath $Path -NoTypeInformation -Encoding UTF8 }
    }

    function Set-DisplayText {
        param([hashtable]$Ui,[string]$BaseName,[string]$Value)
        $display = $Ui["${BaseName}Display"]
        if ($display) { $display.Text = $Value }
        $textbox = $Ui["${BaseName}TextBox"]
        if ($textbox) { $textbox.Text = $Value }
    }

    function Parse-DateLoose {
        param([string]$s)
        if ([string]::IsNullOrWhiteSpace($s)) { return $null }
        try { return [datetime]$s } catch {}
        foreach ($fmt in @('yyyy-MM-dd','dd MMMM yyyy','dd MMM yyyy','MM/dd/yyyy HH:mm:ss','MM/dd/yyyy')) {
            try { return [datetime]::ParseExact($s.Trim(), $fmt, [Globalization.CultureInfo]::InvariantCulture) } catch {}
        }
        return $null
    }

    function Get-PropertyValue {
        param(
            [Parameter(Mandatory=$true)][psobject]$InputObject,
            [Parameter(Mandatory=$true)][string]$PropertyName
        )
        $prop = $InputObject.PSObject.Properties[$PropertyName]
        if ($null -eq $prop) { return $null }
        return $prop.Value
    }

    function Get-TrimmedPropertyValue {
        param(
            [Parameter(Mandatory=$true)][psobject]$InputObject,
            [Parameter(Mandatory=$true)][string]$PropertyName
        )
        return ('' + (Get-PropertyValue -InputObject $InputObject -PropertyName $PropertyName)).Trim()
    }

    function Fmt-DateLong { param($dt) if ($dt) { try { return ([datetime]$dt).ToString('dd MMMM yyyy') } catch {} } return '' }

    function Extract-Ritm {
        param([string]$po)
        if ([string]::IsNullOrWhiteSpace($po)) { return '' }
        if ($po -match '(RITM\d+|TRP\s*-\s*\d{1,2}\s+\w+\s+\d{4})') { return $matches[1] }
        return $po
    }

    function Get-DeviceType {
        param([psobject]$Record)
        if (-not $Record) { return '' }
        $type = ('' + $Record.Type).Trim()
        if ($type -eq 'Computer' -and ('' + $Record.Name).Trim().ToUpper().StartsWith('AO')) { return 'Tangent' }
        return $type
    }

    function Load-DataSet {
        param([string]$ResolvedXamlPath)
        $dataFolder = Get-DataFolder -ResolvedXamlPath $ResolvedXamlPath
        if (-not (Test-Path -LiteralPath $dataFolder)) { throw "Data folder not found: $dataFolder" }

        $records = New-Object System.Collections.Generic.List[object]
        $locationRows = New-Object System.Collections.Generic.List[object]

        foreach ($csvFile in (Get-ChildItem -LiteralPath $dataFolder -Recurse -Filter '*.csv')) {
            $name = $csvFile.Name
            $site = Split-Path -Leaf (Split-Path -Parent $csvFile.FullName)
            $rows = @()
            try { $rows = Import-Csv -LiteralPath $csvFile.FullName } catch { $rows = @() }
            foreach ($r in $rows) {
                if ($name -like 'LocationMaster*') {
                    $locationRows.Add([pscustomobject]@{
                        City = ('' + $site).Trim()
                        Location = Get-TrimmedPropertyValue -InputObject $r -PropertyName 'location'
                        Building = Get-TrimmedPropertyValue -InputObject $r -PropertyName 'u_building'
                        Floor = Get-TrimmedPropertyValue -InputObject $r -PropertyName 'u_floor'
                        Room = Get-TrimmedPropertyValue -InputObject $r -PropertyName 'u_room'
                        Department = Get-TrimmedPropertyValue -InputObject $r -PropertyName 'u_department'
                    }) | Out-Null
                    continue
                }

                $type = $null
                $kind = 'Peripheral'
                $parent = ''
                switch -Wildcard ($name) {
                    'Computers*' { $type = 'Computer'; $kind = 'Computer' }
                    'Monitors*'  { $type = 'Monitor'; $parent = Get-TrimmedPropertyValue -InputObject $r -PropertyName 'u_parent_asset' }
                    'Mics*'      { $type = 'Mic'; $parent = Get-TrimmedPropertyValue -InputObject $r -PropertyName 'u_parent_asset' }
                    'Scanners*'  { $type = 'Scanner'; $parent = Get-TrimmedPropertyValue -InputObject $r -PropertyName 'u_parent_asset' }
                    'Carts*'     { $type = 'Cart'; $parent = Get-TrimmedPropertyValue -InputObject $r -PropertyName 'u_parent_asset' }
                    default { }
                }
                if (-not $type) { continue }

                $assetTagFromTag = Get-PropertyValue -InputObject $r -PropertyName 'asset_tag'
                $assetTagFromAsset = Get-PropertyValue -InputObject $r -PropertyName 'asset'
                $assetTag = if ($assetTagFromTag) { $assetTagFromTag } elseif ($assetTagFromAsset) { $assetTagFromAsset } else { '' }
                $locationCity = Get-PropertyValue -InputObject $r -PropertyName 'location.city'
                $record = [pscustomobject]@{
                    Kind = $kind
                    Type = $type
                    Name = Get-TrimmedPropertyValue -InputObject $r -PropertyName 'name'
                    AssetTag = ('' + $assetTag).Trim()
                    Serial = Get-TrimmedPropertyValue -InputObject $r -PropertyName 'serial_number'
                    Parent = ('' + $parent).Trim()
                    City = if ([string]::IsNullOrWhiteSpace(('' + $locationCity))) { ('' + $site).Trim() } else { ('' + $locationCity).Trim() }
                    Location = Get-TrimmedPropertyValue -InputObject $r -PropertyName 'location'
                    Building = Get-TrimmedPropertyValue -InputObject $r -PropertyName 'u_building'
                    Floor = Get-TrimmedPropertyValue -InputObject $r -PropertyName 'u_floor'
                    Room = Get-TrimmedPropertyValue -InputObject $r -PropertyName 'u_room'
                    Department = Get-TrimmedPropertyValue -InputObject $r -PropertyName 'u_department_location'
                    RITM = (Extract-Ritm -po (Get-TrimmedPropertyValue -InputObject $r -PropertyName 'po_number'))
                    Retire = Parse-DateLoose -s (Get-TrimmedPropertyValue -InputObject $r -PropertyName 'u_scheduled_retirement')
                    LastRounded = Parse-DateLoose -s (Get-TrimmedPropertyValue -InputObject $r -PropertyName 'u_last_rounded_date')
                    MaintenanceType = Get-TrimmedPropertyValue -InputObject $r -PropertyName 'u_device_rounding'
                }
                $records.Add($record) | Out-Null
            }
        }

        $indexByAsset = @{}
        $indexByName = @{}
        $indexBySerial = @{}
        $childrenByParent = @{}
        foreach ($r in $records) {
            if ($r.AssetTag) { $indexByAsset[$r.AssetTag.ToUpper()] = $r }
            if ($r.Name) { $indexByName[$r.Name.ToUpper()] = $r }
            if ($r.Serial) { $indexBySerial[$r.Serial.ToUpper()] = $r }
            if ($r.Parent) {
                $k = $r.Parent.ToUpper()
                if (-not $childrenByParent.ContainsKey($k)) { $childrenByParent[$k] = New-Object System.Collections.Generic.List[object] }
                $childrenByParent[$k].Add($r)
            }
        }

        return [pscustomobject]@{
            Records = @($records.ToArray())
            IndexByAsset = $indexByAsset
            IndexByName = $indexByName
            IndexBySerial = $indexBySerial
            ChildrenByParent = $childrenByParent
            LocationRows = @($locationRows.ToArray())
        }
    }

    function Find-Record {
        param([string]$SearchTerm,[pscustomobject]$DataSet)
        $term = ('' + $SearchTerm).Trim()
        if ([string]::IsNullOrWhiteSpace($term)) { return $null }
        $upper = $term.ToUpper()
        if ($DataSet.IndexByAsset.ContainsKey($upper)) { return $DataSet.IndexByAsset[$upper] }
        if ($DataSet.IndexBySerial.ContainsKey($upper)) { return $DataSet.IndexBySerial[$upper] }
        if ($DataSet.IndexByName.ContainsKey($upper)) { return $DataSet.IndexByName[$upper] }

        foreach ($r in $DataSet.Records) {
            if ($r.AssetTag -like "*$term*" -or $r.Serial -like "*$term*" -or $r.Name -like "*$term*") { return $r }
        }
        return $null
    }

    function Get-AssociatedRows {
        param([pscustomobject]$Current,[pscustomobject]$DataSet)
        $rows = New-Object System.Collections.Generic.List[object]
        if (-not $Current) { return @($rows) }
        $rows.Add([pscustomobject]@{ Role='Parent'; Type=(Get-DeviceType $Current); Name=$Current.Name; AssetTag=$Current.AssetTag; Serial=$Current.Serial; RITM=$Current.RITM; Retire=(Fmt-DateLong $Current.Retire) }) | Out-Null
        if ($Current.AssetTag) {
            $key = $Current.AssetTag.ToUpper()
            if ($DataSet.ChildrenByParent.ContainsKey($key)) {
                foreach ($child in $DataSet.ChildrenByParent[$key]) {
                    $rows.Add([pscustomobject]@{ Role='Child'; Type=(Get-DeviceType $child); Name=$child.Name; AssetTag=$child.AssetTag; Serial=$child.Serial; RITM=$child.RITM; Retire=(Fmt-DateLong $child.Retire) }) | Out-Null
                }
            }
        }
        return @($rows)
    }

    function Get-NearbyRows {
        param([pscustomobject]$Current,[pscustomobject]$DataSet)
        $rows = New-Object System.Collections.Generic.List[object]
        if (-not $Current) { return @($rows) }
        $sameLoc = $DataSet.Records | Where-Object { $_.Type -eq 'Computer' -and $_.Location -eq $Current.Location }
        foreach ($pc in $sameLoc) {
            $days = ''
            if ($pc.LastRounded) {
                try { $days = [math]::Max(0, [int]((New-TimeSpan -Start $pc.LastRounded -End (Get-Date)).TotalDays)) } catch { $days = '' }
            }
            $rows.Add([pscustomobject]@{
                HostName = $pc.Name
                IPAddress = ''
                Subnet = ''
                AssetTag = $pc.AssetTag
                Location = $pc.Location
                Building = $pc.Building
                Floor = $pc.Floor
                Room = $pc.Room
                Department = $pc.Department
                MaintenanceType = if ($pc.MaintenanceType) { $pc.MaintenanceType } else { 'General Rounding' }
                LastRounded = Fmt-DateLong $pc.LastRounded
                DaysAgo = "$days"
                Status = '-'
            }) | Out-Null
        }
        return @($rows)
    }

    function Populate-LocationCombos {
        param([hashtable]$Ui,[psobject]$Current,[pscustomobject]$DataSet)
        $pairs = @(
            @{ Combo=$Ui.CityComboBox; Value=$Current.City; Col='City' },
            @{ Combo=$Ui.LocationComboBox; Value=$Current.Location; Col='Location' },
            @{ Combo=$Ui.BuildingComboBox; Value=$Current.Building; Col='Building' },
            @{ Combo=$Ui.FloorComboBox; Value=$Current.Floor; Col='Floor' },
            @{ Combo=$Ui.RoomComboBox; Value=$Current.Room; Col='Room' },
            @{ Combo=$Ui.DepartmentComboBox; Value=$Current.Department; Col='Department' }
        )
        foreach ($pair in $pairs) {
            $pair.Combo.Items.Clear()
            $seen = New-Object 'System.Collections.Generic.HashSet[string]'
            if ($pair.Value) { [void]$seen.Add([string]$pair.Value); [void]$pair.Combo.Items.Add([string]$pair.Value) }
            foreach ($r in $DataSet.LocationRows) {
                $candidate = ('' + $r.($pair.Col)).Trim()
                if (-not [string]::IsNullOrWhiteSpace($candidate) -and $seen.Add($candidate)) {
                    [void]$pair.Combo.Items.Add($candidate)
                }
            }
            $pair.Combo.Text = [string]$pair.Value
        }
    }

    function Set-WindowDataBindings {
        param([hashtable]$Ui,[pscustomobject]$Current,[pscustomobject]$DataSet)
        if (-not $Current) { return }

        $parentText = if ($Current.Parent) { $Current.Parent } else { '(n/a)' }
        $lastRoundedText = if ($Current.LastRounded) {
            $d = [math]::Max(0, [int]((New-TimeSpan -Start $Current.LastRounded -End (Get-Date)).TotalDays))
            "$(Fmt-DateLong $Current.LastRounded) - $d day$(if($d -eq 1){''} else {'s'}) ago"
        } else { '' }

        $Ui.SearchTextBox.Text = $Current.Name
        $Ui.SelectedDeviceText.Text = $Current.Name
        Set-DisplayText -Ui $Ui -BaseName 'DetectedType' -Value (Get-DeviceType $Current)
        Set-DisplayText -Ui $Ui -BaseName 'HostName' -Value $Current.Name
        Set-DisplayText -Ui $Ui -BaseName 'AssetTag' -Value $Current.AssetTag
        Set-DisplayText -Ui $Ui -BaseName 'Serial' -Value $Current.Serial
        Set-DisplayText -Ui $Ui -BaseName 'Parent' -Value $parentText
        Set-DisplayText -Ui $Ui -BaseName 'Ritm' -Value $Current.RITM
        Set-DisplayText -Ui $Ui -BaseName 'Retire' -Value (Fmt-DateLong $Current.Retire)
        $Ui.LastRoundedText.Text = $lastRoundedText
        $Ui.CityTextBox.Text = $Current.City
        $Ui.LocationTextBox.Text = $Current.Location
        $Ui.BuildingTextBox.Text = $Current.Building
        $Ui.FloorTextBox.Text = $Current.Floor
        $Ui.RoomTextBox.Text = $Current.Room
        $Ui.DepartmentTextBox.Text = $Current.Department
        $Ui.LastQueryBadgeText.Text = "Queried $(Get-Date -Format 'HH:mm')"

        $associated = Get-AssociatedRows -Current $Current -DataSet $DataSet
        $nearby = Get-NearbyRows -Current $Current -DataSet $DataSet

        $Ui.AssociatedDevicesDataGrid.ItemsSource = $associated
        $Ui.NearbyDataGrid.ItemsSource = $nearby
        $Ui.NearbyScopeSummaryText.Text = "Nearby scopes (Location): 1 - Showing $($nearby.Count)"

        Populate-LocationCombos -Ui $Ui -Current $Current -DataSet $DataSet
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
            Timestamp       = (Get-Date).ToString('yyyy-MM-dd HH:mm:ss')
            AssetTag        = $CurrentDevice.AssetTag
            Name            = $CurrentDevice.Name
            Serial          = $CurrentDevice.Serial
            City            = $CurrentDevice.City
            Location        = $CurrentDevice.Location
            Building        = $CurrentDevice.Building
            Floor           = $CurrentDevice.Floor
            Room            = $CurrentDevice.Room
            CheckStatus     = $Ui.CheckStatusComboBox.Text
            RoundingMinutes = (Get-RoundingMinutes -Ui $Ui)
            CableMgmtOK     = if ($Ui.ValidateCableCheckBox.IsChecked) { 'Yes' } else { 'No' }
            CablingNeeded   = if ($Ui.CablingNeededCheckBox.IsChecked) { 'Yes' } else { 'No' }
            LabelOK         = if ($Ui.LabelMonitorCheckBox.IsChecked) { 'Yes' } else { 'No' }
            CartOK          = if ($Ui.PhysicalCartCheckBox.IsChecked) { 'Yes' } else { 'No' }
            PeripheralsOK   = if ($Ui.ValidatePeripheralsCheckBox.IsChecked) { 'Yes' } else { 'No' }
            MaintenanceType = $Ui.MaintenanceTypeComboBox.Text
            Department      = $Ui.DepartmentTextBox.Text
            RoundingUrl     = ''
            Comments        = $Ui.CommentsTextBox.Text
            Rounded         = 'No'
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
                Timestamp       = (Get-Date).ToString('yyyy-MM-dd HH:mm:ss')
                AssetTag        = $item.AssetTag
                Name            = $item.HostName
                Serial          = ''
                City            = $Ui.CityTextBox.Text
                Location        = $item.Location
                Building        = $item.Building
                Floor           = $item.Floor
                Room            = $item.Room
                CheckStatus     = $(if ([string]::IsNullOrWhiteSpace($item.Status) -or $item.Status -eq '-') { 'Complete' } else { $item.Status })
                RoundingMinutes = 3
                CableMgmtOK     = 'Yes'
                CablingNeeded   = 'No'
                LabelOK         = 'Yes'
                CartOK          = 'Yes'
                PeripheralsOK   = 'Yes'
                MaintenanceType = $item.MaintenanceType
                Department      = $item.Department
                RoundingUrl     = ''
                Comments        = 'Saved from Nearby tab'
                Rounded         = 'No'
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

    function Get-NearbySubnetValue {
        param([string]$IpAddress)
        if ([string]::IsNullOrWhiteSpace($IpAddress)) { return '' }
        $ip = $IpAddress.Trim()
        if ($ip.StartsWith('10.64.')) { return 'VPN' }
        return 'Unknown'
    }

    function Invoke-NearbyPingAll {
        param([hashtable]$Ui)
        $rows = @($Ui.NearbyDataGrid.ItemsSource)
        if ($rows.Count -eq 0) {
            Set-StatusMessage -Ui $Ui -Mode 'Warning' -CustomText 'No nearby hosts available to ping'
            return
        }
        foreach ($row in $rows) {
            $hostName = ('' + $row.HostName).Trim()
            if ([string]::IsNullOrWhiteSpace($hostName)) { continue }
            $success = $false
            $ipAddress = ''
            try {
                $ping = New-Object System.Net.NetworkInformation.Ping
                try {
                    $reply = $ping.Send($hostName, 2000)
                    if ($reply -and $reply.Status -eq [System.Net.NetworkInformation.IPStatus]::Success) {
                        $success = $true
                        if ($reply.Address) { $ipAddress = '' + $reply.Address }
                    }
                } finally {
                    $ping.Dispose()
                }
            } catch {
                $success = $false
            }
            $row.IPAddress = $ipAddress
            $row.Subnet = Get-NearbySubnetValue -IpAddress $ipAddress
            $row.Status = if ($success) { 'Complete' } else { 'Inaccessible - Laptop is not available' }
        }
        $Ui.NearbyDataGrid.Items.Refresh()
        Set-StatusMessage -Ui $Ui -Mode 'PingComplete'
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

    $dataSet = Load-DataSet -ResolvedXamlPath $resolvedXamlPath
    $initialDevice = $dataSet.Records | Where-Object { $_.Type -eq 'Computer' } | Select-Object -First 1
    if (-not $initialDevice) { $initialDevice = $dataSet.Records | Select-Object -First 1 }

    $script:AppState = [pscustomobject]@{ LastStatusMode='Found'; DataSet=$dataSet; CurrentDevice=$initialDevice }

    if ($script:AppState.CurrentDevice) {
        Set-WindowDataBindings -Ui $ui -Current $script:AppState.CurrentDevice -DataSet $script:AppState.DataSet
        Set-StatusMessage -Ui $ui -Mode 'Found'
    }
    Toggle-LocationEditMode -Ui $ui -IsEditing:$false

    $ui.DataPathText.Text = "Data: $(Get-DataFolder -ResolvedXamlPath $resolvedXamlPath)"
    $ui.OutputPathText.Text = "Output: $(Get-OutputFolder -ResolvedXamlPath $resolvedXamlPath)"

    $ui.QueryButton.Add_Click({
        $match = Find-Record -SearchTerm $ui.SearchTextBox.Text -DataSet $script:AppState.DataSet
        if ($null -eq $match) {
            Set-StatusMessage -Ui $ui -Mode 'Warning' -CustomText 'No matching device found'
            return
        }
        $script:AppState.CurrentDevice = $match
        Set-WindowDataBindings -Ui $ui -Current $script:AppState.CurrentDevice -DataSet $script:AppState.DataSet
        Set-StatusMessage -Ui $ui -Mode 'Found'
    })

    $ui.SearchTextBox.Add_KeyDown({
        if ($_.Key -eq 'Return') {
            $ui.QueryButton.RaiseEvent((New-Object System.Windows.RoutedEventArgs([System.Windows.Controls.Button]::ClickEvent)))
        }
    })

    $ui.PingButton.Add_Click({
        if (-not $script:AppState.CurrentDevice) { return }
        $host = $script:AppState.CurrentDevice.Name
        $pingOk = $false
        try {
            $ping = New-Object System.Net.NetworkInformation.Ping
            try {
                $reply = $ping.Send($host, 2000)
                $pingOk = ($reply -and $reply.Status -eq [System.Net.NetworkInformation.IPStatus]::Success)
            } finally { $ping.Dispose() }
        } catch {}
        if ($pingOk) {
            $ui.DeviceOnlineText.Text = 'Online'
            Set-StatusMessage -Ui $ui -Mode 'PingComplete' -CustomText 'Device ping successful'
        } else {
            $ui.DeviceOnlineText.Text = 'Offline'
            Set-StatusMessage -Ui $ui -Mode 'Warning' -CustomText 'Device is not pingable'
        }
    })
    $ui.LiveDetailsButton.Add_Click({ [System.Windows.MessageBox]::Show('Phase 2: Live Details migration pending.', 'Live Details') | Out-Null })
    $ui.MonitorLabelButton.Add_Click({ [System.Windows.MessageBox]::Show('Phase 2: Monitor Label migration pending.', 'Monitor Label') | Out-Null })
    $ui.FixNameButton.Add_Click({ [System.Windows.MessageBox]::Show('Phase 2: Fix Name migration pending.', 'Fix Name') | Out-Null })
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
    $ui.CheckCompleteButton.Add_Click({
        $ui.CheckStatusComboBox.Text = 'Complete'
        $ui.ValidateCableCheckBox.IsChecked = $true
        $ui.LabelMonitorCheckBox.IsChecked = $true
        $ui.ValidatePeripheralsCheckBox.IsChecked = $true
        $ui.PhysicalCartCheckBox.IsChecked = $true
    })
    $ui.SaveEventButton.Add_Click({ Save-RoundingEvent -Ui $ui -CurrentDevice $script:AppState.CurrentDevice -ResolvedXamlPath $resolvedXamlPath })
    $ui.ManualRoundButton.Add_Click({ [System.Windows.MessageBox]::Show('Phase 2: Manual Round URL migration pending.', 'Manual Round') | Out-Null })
    $ui.RebuildNearbyButton.Add_Click({
        if ($script:AppState.CurrentDevice) {
            $nearby = Get-NearbyRows -Current $script:AppState.CurrentDevice -DataSet $script:AppState.DataSet
            $ui.NearbyDataGrid.ItemsSource = $nearby
            $ui.NearbyScopeSummaryText.Text = "Nearby scopes (Location): 1 - Showing $($nearby.Count)"
        }
    })
    $ui.ClearNearbyButton.Add_Click({ $ui.NearbyDataGrid.UnselectAll() })
    $ui.PingAllButton.Add_Click({ Invoke-NearbyPingAll -Ui $ui })
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
