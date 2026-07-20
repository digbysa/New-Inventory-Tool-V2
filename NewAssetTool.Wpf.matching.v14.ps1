[CmdletBinding()]
param(
    [string]$XamlPath = $(if ($PSScriptRoot) { Join-Path $PSScriptRoot 'NewAssetTool.matching.v14.xaml' } else { 'NewAssetTool.matching.v3.xaml' })
)

Set-StrictMode -Version 3
$ErrorActionPreference = 'Stop'

Add-Type -AssemblyName PresentationCore
Add-Type -AssemblyName PresentationFramework
Add-Type -AssemblyName WindowsBase
Add-Type -AssemblyName System.Windows.Forms

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

    function Find-ControlByName {
        param(
            [Parameter(Mandatory=$true)][System.Object]$Root,
            [Parameter(Mandatory=$true)][string]$Name
        )

        if ($null -eq $Root) { return $null }

        if ($Root -is [System.Windows.FrameworkElement] -and $Root.Name -eq $Name) {
            return $Root
        }

        if ($Root -is [System.Windows.FrameworkElement]) {
            $named = $Root.FindName($Name)
            if ($null -ne $named) { return $named }
        }

        if ($Root -is [System.Windows.Controls.Control]) {
            $Root.ApplyTemplate()
            if ($Root.Template) {
                $templated = $Root.Template.FindName($Name, $Root)
                if ($null -ne $templated) { return $templated }
            }
        }

        foreach ($child in [System.Windows.LogicalTreeHelper]::GetChildren($Root)) {
            $found = Find-ControlByName -Root $child -Name $Name
            if ($null -ne $found) { return $found }
        }

        if ($Root -is [System.Windows.Media.Visual] -or $Root -is [System.Windows.Media.Media3D.Visual3D]) {
            $count = [System.Windows.Media.VisualTreeHelper]::GetChildrenCount($Root)
            for ($i = 0; $i -lt $count; $i++) {
                $found = Find-ControlByName -Root ([System.Windows.Media.VisualTreeHelper]::GetChild($Root, $i)) -Name $Name
                if ($null -ne $found) { return $found }
            }
        }

        return $null
    }

    function Get-NamedControls {
        param(
            [Parameter(Mandatory=$true)][System.Windows.Window]$Window,
            [Parameter(Mandatory=$true)][string[]]$Names
        )
        $controls = @{}
        foreach ($name in $Names) {
            $control = Find-ControlByName -Root $Window -Name $name
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
    function Get-CmdbUpdatesPath { param([string]$ResolvedXamlPath) Join-Path (Get-OutputFolder -ResolvedXamlPath $ResolvedXamlPath) 'CMDBUpdates.csv' }
    function Get-RoundingMapPath { param([string]$ResolvedXamlPath) Join-Path (Split-Path -Parent $ResolvedXamlPath) 'Data/Rounding.csv' }

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

    $script:RoundingEventColumns = @(
        'Timestamp','AssetTag','Name','Serial','City','Location','Building','Floor','Room',
        'CheckStatus','RoundingMinutes','CableMgmtOK','CablingNeeded','LabelOK','CartOK',
        'PeripheralsOK','MaintenanceType','Department','RoundingUrl','Comments','Rounded'
    )

    function Convert-ToRoundingEventRecord {
        param([psobject]$Row)
        $record = [ordered]@{}
        foreach ($column in $script:RoundingEventColumns) {
            $value = ''
            if ($Row -and ($Row.PSObject.Properties.Name -contains $column)) { $value = $Row.$column }
            $record[$column] = $value
        }
        return [pscustomobject]$record
    }

    function Add-RoundingCsvRow {
        param([string]$Path,[psobject]$Row)
        $rows = @()
        if (Test-Path -LiteralPath $Path) {
            try { $rows = @(Import-Csv -LiteralPath $Path) } catch { $rows = @() }
        }
        $rows += $Row
        $rows | ForEach-Object { Convert-ToRoundingEventRecord -Row $_ } | Export-Csv -LiteralPath $Path -NoTypeInformation
    }

    function Set-RowFieldValue {
        param([object]$Row,[string]$Name,[object]$Value)
        if (-not $Row -or [string]::IsNullOrWhiteSpace($Name)) { return }
        if ($Row.PSObject.Properties.Name -contains $Name) { $Row.$Name = $Value }
        else { $Row | Add-Member -NotePropertyName $Name -NotePropertyValue $Value -Force }
    }

    function Set-DisplayText {
        param([hashtable]$Ui,[string]$BaseName,[string]$Value)
        $display = $Ui["${BaseName}Display"]
        if ($display) { Set-ControlText -Control $display -Value $Value }
        $textbox = $Ui["${BaseName}TextBox"]
        if ($textbox) { Set-ControlText -Control $textbox -Value $Value }
    }

    function Set-ControlText {
        param([object]$Control,[string]$Value)
        if ($null -eq $Control) { return }
        if ($Control.PSObject.Properties.Name -contains 'Text') { $Control.Text = $Value; return }
        if ($Control.PSObject.Properties.Name -contains 'Content') { $Control.Content = $Value; return }
    }

    function Parse-DateLoose {
        param([string]$Value)
        if ([string]::IsNullOrWhiteSpace($Value)) { return $null }
        $formats = @('yyyy-MM-dd','yyyy/MM/dd','MM/dd/yyyy','MM-dd-yyyy','dd/MM/yyyy','dd-MM-yyyy','d/M/yyyy','M/d/yyyy','dd MMMM yyyy','d MMMM yyyy','dd MMM yyyy','d MMM yyyy')
        foreach ($format in $formats) {
            try {
                return [datetime]::ParseExact($Value, $format, [System.Globalization.CultureInfo]::InvariantCulture, [System.Globalization.DateTimeStyles]::AssumeLocal)
            } catch {}
        }
        try { return (Get-Date -Date $Value) } catch { return $null }
    }

    function Format-DateLong {
        param([string]$Value)
        $dt = Parse-DateLoose -Value $Value
        if (-not $dt) { return $Value }
        return $dt.ToString('dd MMMM yyyy')
    }

    function Extract-Ritm {
        param([string]$Value)
        if ([string]::IsNullOrWhiteSpace($Value)) { return '' }
        $trimmed = $Value.Trim()
        $ritm = [regex]::Match($trimmed, '(RITM\d+)')
        if ($ritm.Success) { return $ritm.Groups[1].Value }
        $trpEight = [regex]::Match($trimmed, 'TRP(?<date>\d{8})')
        if ($trpEight.Success) {
            try {
                $dt = [datetime]::ParseExact($trpEight.Groups['date'].Value, 'yyyyMMdd', [System.Globalization.CultureInfo]::InvariantCulture)
                return ('TRP - {0}' -f $dt.ToString('dd MMMM yyyy'))
            } catch {}
        }
        $trpSix = [regex]::Match($trimmed, 'TRP(?<date>\d{6})(?!\d)')
        if ($trpSix.Success) {
            try {
                $digits = $trpSix.Groups['date'].Value
                $month = [int]$digits.Substring(0,2)
                $day = [int]$digits.Substring(2,2)
                $year = 2000 + [int]$digits.Substring(4,2)
                $dt = [datetime]::new($year, $month, $day)
                return ('TRP - {0}' -f $dt.ToString('dd MMMM yyyy'))
            } catch {}
        }
        return $trimmed
    }

    function Get-RoundingStatus {
        param([Nullable[datetime]]$RoundedDate)
        if (-not $RoundedDate) { return 'Red' }
        $daysAgo = [int](((Get-Date).Date - $RoundedDate.Date).TotalDays)
        if ($daysAgo -lt 7) { return 'Green' }
        if ($daysAgo -lt 35) { return 'Yellow' }
        return 'Red'
    }

    function Set-LastRoundedDisplay {
        param([hashtable]$Ui,[string]$LastRoundedRaw)
        $dt = Parse-DateLoose -Value $LastRoundedRaw
        if (-not $dt) {
            Set-ControlText -Control $Ui.LastRoundedText -Value ''
            $Ui.LastRoundedText.Foreground = New-Brush '#BE123C'
            $Ui.LastRoundedContainer.Background = New-Brush '#FDF0F1'
            $Ui.LastRoundedContainer.BorderBrush = New-Brush '#F5C2C7'
            return
        }
        $dateText = $dt.ToString('dd MMMM yyyy')
        $daysAgo = [int](((Get-Date).Date - $dt.Date).TotalDays)
        if ($daysAgo -le 0) { $Ui.LastRoundedText.Text = "$dateText - Today" }
        else {
            $plural = if ($daysAgo -eq 1) { '' } else { 's' }
            $Ui.LastRoundedText.Text = ("{0} - {1} day{2} ago" -f $dateText, $daysAgo, $plural)
        }
        $status = Get-RoundingStatus -RoundedDate $dt
        switch ($status) {
            'Green' {
                $Ui.LastRoundedText.Foreground = New-Brush '#15803D'
                $Ui.LastRoundedLabelText.Foreground = New-Brush '#15803D'
                $Ui.LastRoundedContainer.Background = New-Brush '#ECFDF3'
                $Ui.LastRoundedContainer.BorderBrush = New-Brush '#BBF7D0'
                $Ui.LastRoundedAttentionBadge.Visibility = 'Collapsed'
            }
            'Yellow' {
                $Ui.LastRoundedText.Foreground = New-Brush '#B45309'
                $Ui.LastRoundedLabelText.Foreground = New-Brush '#B45309'
                $Ui.LastRoundedContainer.Background = New-Brush '#FEE7C3'
                $Ui.LastRoundedContainer.BorderBrush = New-Brush '#FCD49B'
                $Ui.LastRoundedAttentionBadge.Visibility = 'Collapsed'
            }
            default {
                $Ui.LastRoundedText.Foreground = New-Brush '#BE123C'
                $Ui.LastRoundedLabelText.Foreground = New-Brush '#BE123C'
                $Ui.LastRoundedContainer.Background = New-Brush '#FDF0F1'
                $Ui.LastRoundedContainer.BorderBrush = New-Brush '#F5C2C7'
                $Ui.LastRoundedAttentionBadge.Visibility = 'Visible'
                $Ui.LastRoundedAttentionText.Text = 'Attention needed'
            }
        }
    }


    function Get-FieldValue {
        param([object]$Row,[string[]]$Names)
        if (-not $Row) { return '' }
        foreach ($name in $Names) {
            $property = $Row.PSObject.Properties[$name]
            if ($property) {
                $value = $property.Value
                if (-not [string]::IsNullOrWhiteSpace([string]$value)) { return [string]$value }
            }
        }
        return ''
    }

    function Import-InventoryCsv {
        param([string]$Path)
        try { return @(Import-Csv -LiteralPath $Path -ErrorAction Stop) } catch { return @() }
    }

    function Get-InventorySites {
        param([string]$DataRoot)
        if (-not (Test-Path -LiteralPath $DataRoot)) { return @() }
        $sites = @()
        foreach ($dir in Get-ChildItem -LiteralPath $DataRoot -Directory) {
            $hasComputers = @(Get-ChildItem -LiteralPath $dir.FullName -File -Filter 'Computers - *.csv' -ErrorAction SilentlyContinue).Count -gt 0
            $hasMonitors = @(Get-ChildItem -LiteralPath $dir.FullName -File -Filter 'Monitors - *.csv' -ErrorAction SilentlyContinue).Count -gt 0
            if ($hasComputers -or $hasMonitors) {
                $sites += [pscustomobject]@{ Name = $dir.Name; FolderPath = $dir.FullName }
            }
        }
        return @($sites | Sort-Object Name)
    }

    function Select-InventorySite {
        param([object[]]$Sites)
        if (-not $Sites -or $Sites.Count -eq 0) { return $null }
        if ($Sites.Count -eq 1) { return $Sites[0] }

        $dialog = New-Object System.Windows.Forms.Form
        $dialog.Text = 'Choose your site'
        $dialog.Width = 420
        $dialog.Height = 170
        $dialog.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::FixedDialog
        $dialog.StartPosition = [System.Windows.Forms.FormStartPosition]::CenterScreen
        $dialog.MinimizeBox = $false
        $dialog.MaximizeBox = $false
        $dialog.TopMost = $true

        $label = New-Object System.Windows.Forms.Label
        $label.Text = 'Choose your site:'
        $label.AutoSize = $true
        $label.Left = 16
        $label.Top = 20

        $combo = New-Object System.Windows.Forms.ComboBox
        $combo.Left = 16
        $combo.Top = 46
        $combo.Width = 370
        $combo.DropDownStyle = [System.Windows.Forms.ComboBoxStyle]::DropDownList
        [void]$combo.Items.AddRange(@($Sites.Name))
        $combo.SelectedIndex = 0

        $ok = New-Object System.Windows.Forms.Button
        $ok.Text = 'OK'
        $ok.Left = 230
        $ok.Top = 84
        $ok.Width = 75
        $ok.DialogResult = [System.Windows.Forms.DialogResult]::OK
        $dialog.AcceptButton = $ok

        $cancel = New-Object System.Windows.Forms.Button
        $cancel.Text = 'Cancel'
        $cancel.Left = 311
        $cancel.Top = 84
        $cancel.Width = 75
        $cancel.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
        $dialog.CancelButton = $cancel

        $dialog.Controls.AddRange(@($label,$combo,$ok,$cancel))
        $result = $dialog.ShowDialog()
        if ($result -ne [System.Windows.Forms.DialogResult]::OK -or $combo.SelectedIndex -lt 0) { return $null }
        return $Sites[$combo.SelectedIndex]
    }

    function Load-InventoryData {
        param([string]$ResolvedXamlPath,[string]$SiteFolderPath)
        $dataRoot = Join-Path (Split-Path -Parent $ResolvedXamlPath) 'Data'
        $computers = @(); $monitors = @(); $locations = @(); $carts = @(); $mics = @(); $scanners = @()
        $pathsToLoad = @($dataRoot)
        if (-not [string]::IsNullOrWhiteSpace($SiteFolderPath) -and (Test-Path -LiteralPath $SiteFolderPath)) {
            $pathsToLoad = @($SiteFolderPath)
        }
        if (Test-Path -LiteralPath $dataRoot) {
            foreach ($file in Get-ChildItem -LiteralPath $dataRoot -File -Filter '*.csv') {
                if ($file.Name -eq 'Carts.csv') { $carts += Import-InventoryCsv -Path $file.FullName }
                elseif ($file.Name -eq 'Mics.csv') { $mics += Import-InventoryCsv -Path $file.FullName }
                elseif ($file.Name -eq 'Scanners.csv') { $scanners += Import-InventoryCsv -Path $file.FullName }
            }
            foreach ($path in $pathsToLoad) {
                foreach ($file in Get-ChildItem -LiteralPath $path -Recurse -File -Filter '*.csv') {
                    if ($file.Name -like 'Computers - *.csv') { $computers += Import-InventoryCsv -Path $file.FullName }
                    elseif ($file.Name -like 'Monitors - *.csv') { $monitors += Import-InventoryCsv -Path $file.FullName }
                    elseif ($file.Name -like 'LocationMaster*.csv') { $locations += Import-InventoryCsv -Path $file.FullName }
                }
            }
        }
        $inventory = [pscustomobject]@{
            DataRoot=$dataRoot; SiteFolderPath=$SiteFolderPath
            Computers=$computers; Monitors=$monitors; Carts=$carts; Mics=$mics; Scanners=$scanners; Locations=$locations
        }
        Build-InventoryIndices -Inventory $inventory
        return $inventory
    }

    function ConvertTo-DeviceRecord {
        param([object]$Row,[string]$DetectedType='Computer')
        $name = Get-FieldValue -Row $Row -Names @('name','HostName')
        $parent = Get-FieldValue -Row $Row -Names @('u_parent_asset','Parent')
        if ([string]::IsNullOrWhiteSpace($parent)) { $parent = '(n/a)' }
        return [pscustomobject]@{
            SearchKeys=@($name,(Get-FieldValue -Row $Row -Names @('asset_tag','AssetTag')),(Get-FieldValue -Row $Row -Names @('serial_number','Serial')))
            DetectedType=$DetectedType; Name=$name
            AssetTag=Get-FieldValue -Row $Row -Names @('asset_tag','AssetTag')
            Serial=Get-FieldValue -Row $Row -Names @('serial_number','Serial')
            Parent=$parent; RITM=Extract-Ritm (Get-FieldValue -Row $Row -Names @('po_number','RITM'))
            RetireDate=Format-DateLong (Get-FieldValue -Row $Row -Names @('u_scheduled_retirement','RetireDate'))
            LastRounded=Get-FieldValue -Row $Row -Names @('u_last_rounded_date','LastRounded')
            City=Get-FieldValue -Row $Row -Names @('location.city','City')
            Location=Get-FieldValue -Row $Row -Names @('location','Location')
            Building=Get-FieldValue -Row $Row -Names @('u_building','Building')
            Floor=Get-FieldValue -Row $Row -Names @('u_floor','Floor')
            Room=Get-FieldValue -Row $Row -Names @('u_room','Room')
            Department=Get-FieldValue -Row $Row -Names @('u_department_location','Department')
        }
    }

    function Add-IndexKey {
        param([hashtable]$Index,[string]$Key,[pscustomobject]$Value)
        if (-not $Index -or -not $Value) { return }
        if ([string]::IsNullOrWhiteSpace($Key)) { return }
        $normalized = $Key.Trim().ToUpper()
        if ([string]::IsNullOrWhiteSpace($normalized)) { return }
        $Index[$normalized] = $Value
        $compact = ($normalized -replace '[-\s]','')
        if (-not [string]::IsNullOrWhiteSpace($compact)) { $Index[$compact] = $Value }
    }

    function Build-InventoryIndices {
        param([pscustomobject]$Inventory)
        if (-not $Inventory) { return }
        $indexByName = @{}
        $indexByAsset = @{}
        $indexBySerial = @{}

        foreach ($set in @(
            @{Rows=$Inventory.Computers; Type='Computer'},
            @{Rows=$Inventory.Monitors; Type='Monitor'},
            @{Rows=$Inventory.Carts; Type='Cart'},
            @{Rows=$Inventory.Mics; Type='Mic'},
            @{Rows=$Inventory.Scanners; Type='Scanner'}
        )) {
            foreach ($row in $set.Rows) {
                $record = ConvertTo-DeviceRecord -Row $row -DetectedType $set.Type
                Add-IndexKey -Index $indexByName -Key $record.Name -Value $record
                Add-IndexKey -Index $indexByAsset -Key $record.AssetTag -Value $record
                Add-IndexKey -Index $indexBySerial -Key $record.Serial -Value $record
            }
        }

        $Inventory | Add-Member -NotePropertyName IndexByName -NotePropertyValue $indexByName -Force
        $Inventory | Add-Member -NotePropertyName IndexByAsset -NotePropertyValue $indexByAsset -Force
        $Inventory | Add-Member -NotePropertyName IndexBySerial -NotePropertyValue $indexBySerial -Force
    }

    function Find-InventoryMatch {
        param([string]$SearchTerm,[pscustomobject]$Inventory)
        $term = $SearchTerm.Trim().ToUpper()
        if ([string]::IsNullOrWhiteSpace($term)) { return $null }

        $isComputerName = $term -match '^(PC|LD|TD|AO|WT)'
        $isAssetTag = $term -match '^(HSS|C)'
        $isHssAssetTag = $term -match '^HSS-'

        if ($isComputerName) {
            $match = $Inventory.IndexByName[$term]
            if ($match) { return $match }
        }

        if ($isAssetTag) {
            if ($isHssAssetTag) {
                $match = $Inventory.IndexByAsset[$term]
                if ($match -and $match.DetectedType -in @('Monitor','Computer','Cart')) { return $match }
            }
            else {
                $match = $Inventory.IndexByAsset[$term]
                if ($match) { return $match }
            }
        }

        $match = $Inventory.IndexByName[$term]
        if ($match) { return $match }
        $match = $Inventory.IndexByAsset[$term]
        if ($match) { return $match }
        $match = $Inventory.IndexBySerial[$term]
        if ($match) { return $match }

        $compactTerm = ($term -replace '[-\s]','')
        if ($compactTerm -and $compactTerm -ne $term) {
            $match = $Inventory.IndexByName[$compactTerm]
            if ($match) { return $match }
            $match = $Inventory.IndexByAsset[$compactTerm]
            if ($match) { return $match }
            $match = $Inventory.IndexBySerial[$compactTerm]
            if ($match) { return $match }
        }
        return $null
    }


    function Resolve-ParentDevice {
        param([pscustomobject]$Device,[pscustomobject]$Inventory)
        if (-not $Device) { return $null }
        if ($Device.DetectedType -eq 'Computer' -and $Device.Name -match '^(LD|PC|TD|AO|WT)') { return $Device }

        $tokens = @(@($Device.Parent) | Where-Object { -not [string]::IsNullOrWhiteSpace($_) } | ForEach-Object { $_.Trim().ToUpper() })
        if (-not $tokens -or $tokens.Count -eq 0) { return $null }

        foreach ($row in $Inventory.Computers) {
            $record = ConvertTo-DeviceRecord -Row $row -DetectedType 'Computer'
            $name = if ($record.Name) { $record.Name.Trim().ToUpper() } else { '' }
            if ($name -notmatch '^(LD|PC|TD|AO|WT)') { continue }
            $candidates = @($record.Name,$record.AssetTag,$record.Serial) | Where-Object { -not [string]::IsNullOrWhiteSpace($_) } | ForEach-Object { $_.Trim().ToUpper() }
            foreach ($token in $tokens) {
                if ($candidates -contains $token) { return $record }
            }
        }
        return $null
    }
    function Get-AssociationTokenVariants {
        param([string]$Token)
        $variants = New-Object System.Collections.ArrayList
        if ([string]::IsNullOrWhiteSpace($Token)) { return ,@() }
        $upper = $Token.Trim().ToUpper()
        if ([string]::IsNullOrWhiteSpace($upper)) { return ,@() }
        [void]$variants.Add($upper)
        $compact = ($upper -replace '[-\s]','')
        if (-not [string]::IsNullOrWhiteSpace($compact) -and -not $variants.Contains($compact)) { [void]$variants.Add($compact) }
        $baseHost = ($upper -replace '-MIC$','' -replace '-SCN$','')
        if (-not [string]::IsNullOrWhiteSpace($baseHost) -and -not $variants.Contains($baseHost)) { [void]$variants.Add($baseHost) }
        $baseCompact = ($baseHost -replace '[-\s]','')
        if (-not [string]::IsNullOrWhiteSpace($baseCompact) -and -not $variants.Contains($baseCompact)) { [void]$variants.Add($baseCompact) }
        return ,$variants
    }
    function Get-ComputerTypeFromName {
        param([string]$Name,[string]$FallbackType='Computer')
        if ([string]::IsNullOrWhiteSpace($Name)) { return $FallbackType }
        switch -Regex ($Name.Trim()) {
            '(?i)^LD' { return 'Laptop' }
            '(?i)^PC' { return 'Desktop' }
            '(?i)^AO' { return 'Tangent' }
            '(?i)^WT' { return 'Thin Client' }
            '(?i)^TD' { return 'Tablet' }
            default { return $FallbackType }
        }
    }

    function Get-AssociatedDeviceDisplayType {
        param([pscustomobject]$Device)
        if (-not $Device) { return '' }
        if ($Device.DetectedType -eq 'Computer') {
            return Get-ComputerTypeFromName -Name $Device.Name -FallbackType $Device.DetectedType
        }
        return $Device.DetectedType
    }

    function Build-AssociatedDevices {
        param([pscustomobject]$Device,[pscustomobject]$Inventory)
        $parentDevice = Resolve-ParentDevice -Device $Device -Inventory $Inventory
        $effectiveParent = if ($parentDevice) { $parentDevice } else { $Device }
        $queryRole = if ($parentDevice) { 'Child' } else { 'Parent' }

        $parentType = Get-AssociatedDeviceDisplayType -Device $effectiveParent
        $results = @([pscustomobject]@{ Role='Parent'; Type=$parentType; Name=$effectiveParent.Name; AssetTag=$effectiveParent.AssetTag; Serial=$effectiveParent.Serial; SerialForeground='#1F2937'; SerialToolTip=''; RITM=$effectiveParent.RITM; Retire=(Format-DateLong $effectiveParent.RetireDate); CmdbUrl=(Get-CmdbLink -DeviceType $effectiveParent.DetectedType -AssetTag $effectiveParent.AssetTag); Device=$effectiveParent })
        $childrenByParent = @{}
        foreach ($collectionName in @('Monitors','Carts','Mics','Scanners')) {
            $collection = $Inventory.$collectionName
            if (-not $collection) { continue }
            foreach ($row in $collection) {
                $token = Get-FieldValue -Row $row -Names @('u_parent_asset','Parent')
                if ([string]::IsNullOrWhiteSpace($token)) { continue }
                foreach ($key in (Get-AssociationTokenVariants -Token $token)) {
                    if (-not $childrenByParent.ContainsKey($key)) { $childrenByParent[$key] = @() }
                    $childrenByParent[$key] += ,$row
                }
            }
        }

        $parentTokens = New-Object System.Collections.ArrayList
        foreach ($candidate in @($effectiveParent.AssetTag,$effectiveParent.Name,$effectiveParent.Serial,$Device.Parent)) {
            foreach ($variant in (Get-AssociationTokenVariants -Token $candidate)) {
                if (-not $parentTokens.Contains($variant)) { [void]$parentTokens.Add($variant) }
            }
        }

        $addedChildAssetTags = @{}
        foreach ($token in $parentTokens) {
            if (-not $childrenByParent.ContainsKey($token)) { continue }
            foreach ($row in $childrenByParent[$token]) {
                $childAssetTag = (Get-FieldValue -Row $row -Names @('asset_tag')).Trim().ToUpper()
                if (-not [string]::IsNullOrWhiteSpace($childAssetTag) -and $addedChildAssetTags.ContainsKey($childAssetTag)) { continue }
                $type = if ($Inventory.Carts -contains $row) { 'Cart' } elseif ($Inventory.Mics -contains $row) { 'Mic' } elseif ($Inventory.Scanners -contains $row) { 'Scanner' } else { 'Monitor' }
                $role = if ($Device.AssetTag -and ($childAssetTag -eq $Device.AssetTag.Trim().ToUpper())) { $queryRole } else { 'Child' }
                $childDevice = ConvertTo-DeviceRecord -Row $row -DetectedType $type
                $record = [pscustomobject]@{ Role=$role; Type=$type; Name=$childDevice.Name; AssetTag=$childDevice.AssetTag; Serial=$childDevice.Serial; SerialForeground='#1F2937'; SerialToolTip=''; RITM=$childDevice.RITM; Retire=(Format-DateLong $childDevice.RetireDate); CmdbUrl=(Get-CmdbLink -DeviceType $type -AssetTag $childDevice.AssetTag); Device=$childDevice }
                $results += $record
                if (-not [string]::IsNullOrWhiteSpace($childAssetTag)) { $addedChildAssetTags[$childAssetTag] = $true }
            }
        }
        return ,$results
    }

    function Normalize-AssocSearch {
        param([string]$Raw)
        if ([string]::IsNullOrWhiteSpace($Raw)) { return $null }
        return (($Raw.Trim().ToUpper() -replace '\s','') -replace '-','')
    }

    function Resolve-AssociatedPeripheralLookup {
        param([string]$Query,[pscustomobject]$Inventory)
        $normalized = Normalize-AssocSearch -Raw $Query
        $rawUpper = if ([string]::IsNullOrWhiteSpace($Query)) { $null } else { $Query.Trim().ToUpper() }
        if ([string]::IsNullOrWhiteSpace($normalized) -and [string]::IsNullOrWhiteSpace($rawUpper)) {
            return [pscustomobject]@{ NormalizedInput=$null; Candidate=$null }
        }

        $keys = New-Object System.Collections.ArrayList
        $addKey = {
            param([string]$Candidate)
            if ([string]::IsNullOrWhiteSpace($Candidate)) { return }
            $u = $Candidate.Trim().ToUpper()
            if (-not [string]::IsNullOrWhiteSpace($u) -and -not $keys.Contains($u)) { [void]$keys.Add($u) }
            $compact = ($u -replace '[-\s]','')
            if (-not [string]::IsNullOrWhiteSpace($compact) -and -not $keys.Contains($compact)) { [void]$keys.Add($compact) }
        }
        & $addKey $normalized
        & $addKey $rawUpper

        foreach ($set in @(
            @{ Rows=$Inventory.Monitors; Type='Monitor' },
            @{ Rows=$Inventory.Carts; Type='Cart' },
            @{ Rows=$Inventory.Mics; Type='Mic' },
            @{ Rows=$Inventory.Scanners; Type='Scanner' }
        )) {
            foreach ($row in $set.Rows) {
                $record = ConvertTo-DeviceRecord -Row $row -DetectedType $set.Type
                $candidates = @($record.Name,$record.AssetTag,$record.Serial) | Where-Object { -not [string]::IsNullOrWhiteSpace($_) }
                foreach ($candidate in $candidates) {
                    $candKey = (($candidate.Trim().ToUpper() -replace '\s','') -replace '-','')
                    if ($keys -contains $candKey -or $keys -contains $candidate.Trim().ToUpper()) {
                        $record | Add-Member -NotePropertyName Kind -NotePropertyValue 'Peripheral' -Force
                        return [pscustomobject]@{ NormalizedInput=if($normalized){$normalized}else{$rawUpper}; Candidate=$record }
                    }
                }
            }
        }
        return [pscustomobject]@{ NormalizedInput=if($normalized){$normalized}else{$rawUpper}; Candidate=$null }
    }

    function Get-ProposedPeripheralName {
        param([pscustomobject]$Candidate,[pscustomobject]$ParentDevice)
        if (-not $Candidate -or -not $ParentDevice) { return $null }
        switch ($Candidate.DetectedType) {
            'Monitor' { return $ParentDevice.Name }
            'Mic' { return "$($ParentDevice.Name)-Mic" }
            'Scanner' { return "$($ParentDevice.Name)-SCN" }
            'Cart' { return "$($ParentDevice.Name)-CRT" }
            default { return $Candidate.Name }
        }
    }

    function Get-ParentPreviewDisplay {
        param([string]$Token,[pscustomobject]$FallbackParent,[pscustomobject]$Inventory)
        if ([string]::IsNullOrWhiteSpace($Token)) { return '(none)' }
        $trimmed = $Token.Trim()
        if ($FallbackParent -and $FallbackParent.AssetTag -and $trimmed.ToUpper() -eq $FallbackParent.AssetTag.Trim().ToUpper()) { return $FallbackParent.Name }
        if ($Inventory) {
            foreach ($key in (Get-AssociationTokenVariants -Token $trimmed)) {
                if ($Inventory.IndexByAsset -and $Inventory.IndexByAsset.ContainsKey($key)) { return $Inventory.IndexByAsset[$key].Name }
                if ($Inventory.IndexByName -and $Inventory.IndexByName.ContainsKey($key)) { return $Inventory.IndexByName[$key].Name }
            }
        }
        return $trimmed
    }

    function Resolve-ParentAssetTag {
        param([string]$Token,[pscustomobject]$Inventory)
        if ([string]::IsNullOrWhiteSpace($Token)) { return '' }
        $trimmed = $Token.Trim()
        if (-not $Inventory) { return $trimmed }
        foreach ($key in (Get-AssociationTokenVariants -Token $trimmed)) {
            if ($Inventory.IndexByAsset -and $Inventory.IndexByAsset.ContainsKey($key)) { return $Inventory.IndexByAsset[$key].AssetTag }
            if ($Inventory.IndexByName -and $Inventory.IndexByName.ContainsKey($key)) { return $Inventory.IndexByName[$key].AssetTag }
        }
        return $trimmed
    }

    function Add-CmdbAssociationUpdate {
        param(
            [string]$ResolvedXamlPath,
            [pscustomobject]$Candidate,
            [string]$OldParent,
            [string]$NewParent,
            [string]$OldName,
            [string]$NewName,
            [string]$Action = 'Link',
            [pscustomobject]$Inventory
        )
        Ensure-OutputFolder -ResolvedXamlPath $ResolvedXamlPath
        $cmdbLink = Get-CmdbLink -DeviceType $Candidate.DetectedType -AssetTag $Candidate.AssetTag
        $cmdbHyperlink = if ([string]::IsNullOrWhiteSpace($cmdbLink)) { '' } else { "=HYPERLINK(`"$cmdbLink`",`"$($Candidate.AssetTag)`")" }
        Add-CsvRow -Path (Get-CmdbUpdatesPath -ResolvedXamlPath $ResolvedXamlPath) -Row ([pscustomobject]@{
            Timestamp = (Get-Date).ToString('yyyy-MM-dd HH:mm:ss')
            Action = $Action
            DeviceType = $Candidate.DetectedType
            AssetTag = $Candidate.AssetTag
            OldParent = Resolve-ParentAssetTag -Token $OldParent -Inventory $Inventory
            NewParent = Resolve-ParentAssetTag -Token $NewParent -Inventory $Inventory
            OldName = $OldName
            NewName = $NewName
            CMDBLink = $cmdbHyperlink
        })
    }

    function Convert-WmiIdToString {
        param([UInt16[]]$Id)
        if (-not $Id) { return $null }
        $chars = @()
        foreach ($code in $Id) {
            if ($code -le 0) { break }
            if ($code -gt 0 -and $code -lt 256) { $chars += [char]$code }
        }
        if ($chars.Count -eq 0) { return $null }
        return (-join $chars).Trim()
    }

    function Test-ComputerPingable {
        param([string]$ComputerName)
        if ([string]::IsNullOrWhiteSpace($ComputerName)) { return $false }
        try { return [bool](Test-Connection -ComputerName $ComputerName -Count 1 -Quiet -ErrorAction SilentlyContinue) } catch { return $false }
    }

    function ConvertTo-IPv4Bytes {
        param([string]$IpAddress)
        if ([string]::IsNullOrWhiteSpace($IpAddress)) { return $null }
        try {
            $parsed = [System.Net.IPAddress]::Parse($IpAddress.Trim())
            $bytes = $parsed.GetAddressBytes()
            if ($bytes.Length -ne 4) { return $null }
            return $bytes
        } catch {
            return $null
        }
    }

    function Resolve-SubnetName {
        param([string]$IpAddress,[string]$DataRoot)
        if ([string]::IsNullOrWhiteSpace($IpAddress)) { return 'Unknown' }
        $ip = $IpAddress.Trim()
        if ($ip -match '^10\.64\.') { return 'VPN' }

        $subnetPath = Join-Path $DataRoot 'SiteSubnets.csv'
        if (-not (Test-Path -LiteralPath $subnetPath)) { return 'Unknown' }

        $ipBytes = ConvertTo-IPv4Bytes -IpAddress $ip
        if (-not $ipBytes) { return 'Unknown' }

        foreach ($line in (Get-Content -LiteralPath $subnetPath)) {
            if ([string]::IsNullOrWhiteSpace($line)) { continue }
            $parts = $line.Split(',')
            if ($parts.Count -lt 2) { continue }
            $cidr = $parts[0].Trim()
            $subnetName = $parts[1].Trim()
            if ([string]::IsNullOrWhiteSpace($cidr) -or [string]::IsNullOrWhiteSpace($subnetName)) { continue }
            $cidrParts = $cidr.Split('/')
            if ($cidrParts.Count -ne 2) { continue }

            $networkBytes = ConvertTo-IPv4Bytes -IpAddress $cidrParts[0]
            if (-not $networkBytes) { continue }
            try { $prefix = [int]$cidrParts[1] } catch { continue }
            if ($prefix -lt 0 -or $prefix -gt 32) { continue }

            $matches = $true
            for ($i = 0; $i -lt 4; $i++) {
                $bits = [Math]::Min([Math]::Max($prefix - (8 * $i), 0), 8)
                $mask = if ($bits -ge 8) { [byte]255 } elseif ($bits -le 0) { [byte]0 } else { [byte]((0xFF -shl (8 - $bits)) -band 0xFF) }
                if (($ipBytes[$i] -band $mask) -ne ($networkBytes[$i] -band $mask)) { $matches = $false; break }
            }
            if ($matches) { return $subnetName }
        }
        return 'Unknown'
    }

    function Get-IPv4AddressFromPingReply {
        param([object]$Reply,[string]$ComputerName)
        foreach ($name in @('IPV4Address','ProtocolAddress','Address')) {
            $value = Get-FieldValue -Row $Reply -Names @($name)
            if ($value -match '^\d{1,3}(\.\d{1,3}){3}$') { return $value }
        }
        try {
            $addresses = [System.Net.Dns]::GetHostAddresses($ComputerName) | Where-Object { $_.AddressFamily -eq [System.Net.Sockets.AddressFamily]::InterNetwork }
            if ($addresses) { return [string]$addresses[0] }
        } catch {}
        return ''
    }

    function Start-ContinuousPingWindow {
        param([string]$Target)
        if ([string]::IsNullOrWhiteSpace($Target)) { return }
        try {
            if ([Environment]::OSVersion.Platform -eq [System.PlatformID]::Win32NT) {
                Start-Process -FilePath 'cmd.exe' -ArgumentList @('/k','ping','-t',$Target) | Out-Null
            }
        } catch {}
    }

    function Invoke-DevicePing {
        param([string]$ComputerName,[string]$DataRoot)
        if ([string]::IsNullOrWhiteSpace($ComputerName)) { throw 'Enter or query a device before using Ping.' }

        $reply = $null
        try {
            $results = @(Test-Connection -ComputerName $ComputerName -Count 1 -ErrorAction Stop)
            if ($results.Count -gt 0) { $reply = $results[0] }
        } catch {
            return [pscustomobject]@{ HostName=$ComputerName; Success=$false; IpAddress='Unknown'; ResponseTime='Timed out'; Subnet='Unknown'; ErrorMessage=$_.Exception.Message }
        }

        $ipAddress = Get-IPv4AddressFromPingReply -Reply $reply -ComputerName $ComputerName
        if ([string]::IsNullOrWhiteSpace($ipAddress)) { $ipAddress = 'Unknown' }

        $responseMs = Get-FieldValue -Row $reply -Names @('ResponseTime','Latency')
        if ([string]::IsNullOrWhiteSpace($responseMs)) { $responseMs = 'Unknown' } else { $responseMs = "$responseMs ms" }

        return [pscustomobject]@{ HostName=$ComputerName; Success=$true; IpAddress=$ipAddress; ResponseTime=$responseMs; Subnet=(Resolve-SubnetName -IpAddress $ipAddress -DataRoot $DataRoot); ErrorMessage='' }
    }

    function Show-PingResultDialog {
        param([object]$PingResult)
        $message = "Device: $($PingResult.HostName)`nIP Address: $($PingResult.IpAddress)`nPing Time: $($PingResult.ResponseTime)`nSubnet: $($PingResult.Subnet)"
        if (-not $PingResult.Success -and -not [string]::IsNullOrWhiteSpace($PingResult.ErrorMessage)) { $message += "`n`nError: $($PingResult.ErrorMessage)" }
        [System.Windows.MessageBox]::Show($message, 'Ping Result') | Out-Null
    }


    function Show-MonitorLabelDialog {
        param([hashtable]$Ui,[pscustomobject]$ParentDevice)
        if (-not $ParentDevice) { return }

        $assetTag = ''
        $hostName = ''
        try { if ($ParentDevice.AssetTag) { $assetTag = ('' + $ParentDevice.AssetTag).Trim() } } catch {}
        try { if ($ParentDevice.Name) { $hostName = ('' + $ParentDevice.Name).Trim() } } catch {}
        if ([string]::IsNullOrWhiteSpace($assetTag)) { $assetTag = '(blank)' }
        if ([string]::IsNullOrWhiteSpace($hostName)) { $hostName = '(blank)' }

        $window = New-Object System.Windows.Window
        $window.Title = 'Monitor Label'
        $window.SizeToContent = 'WidthAndHeight'
        $window.WindowStartupLocation = 'CenterOwner'
        $window.ResizeMode = 'NoResize'
        $window.ShowInTaskbar = $false
        $window.Owner = [System.Windows.Window]::GetWindow($Ui.MainTabControl)

        $panel = New-Object System.Windows.Controls.StackPanel
        $panel.Margin = '16'
        $window.Content = $panel

        $grid = New-Object System.Windows.Controls.Grid
        $grid.Margin = '0,0,0,18'
        $grid.ColumnDefinitions.Add((New-Object System.Windows.Controls.ColumnDefinition -Property @{ Width='Auto' })) | Out-Null
        $grid.ColumnDefinitions.Add((New-Object System.Windows.Controls.ColumnDefinition -Property @{ Width='Auto' })) | Out-Null
        $grid.RowDefinitions.Add((New-Object System.Windows.Controls.RowDefinition -Property @{ Height='Auto' })) | Out-Null
        $grid.RowDefinitions.Add((New-Object System.Windows.Controls.RowDefinition -Property @{ Height='Auto' })) | Out-Null
        $panel.Children.Add($grid) | Out-Null

        $labelFontSize = 30
        $valueFontSize = 30
        $labelMargin = New-Object System.Windows.Thickness(0,0,14,10)
        $valueMargin = New-Object System.Windows.Thickness(0,0,0,10)

        $assetLabel = New-Object System.Windows.Controls.TextBlock -Property @{ Text='Asset:'; FontSize=$labelFontSize; Margin=$labelMargin; VerticalAlignment='Center' }
        $assetValue = New-Object System.Windows.Controls.TextBlock -Property @{ Text=$assetTag; FontSize=$valueFontSize; FontWeight='Bold'; Margin=$valueMargin; VerticalAlignment='Center' }
        $hostLabel = New-Object System.Windows.Controls.TextBlock -Property @{ Text='Hostname:'; FontSize=$labelFontSize; Margin=$labelMargin; VerticalAlignment='Center' }
        $hostValue = New-Object System.Windows.Controls.TextBlock -Property @{ Text=$hostName; FontSize=$valueFontSize; FontWeight='Bold'; Margin=$valueMargin; VerticalAlignment='Center' }

        [System.Windows.Controls.Grid]::SetRow($assetLabel, 0); [System.Windows.Controls.Grid]::SetColumn($assetLabel, 0)
        [System.Windows.Controls.Grid]::SetRow($assetValue, 0); [System.Windows.Controls.Grid]::SetColumn($assetValue, 1)
        [System.Windows.Controls.Grid]::SetRow($hostLabel, 1); [System.Windows.Controls.Grid]::SetColumn($hostLabel, 0)
        [System.Windows.Controls.Grid]::SetRow($hostValue, 1); [System.Windows.Controls.Grid]::SetColumn($hostValue, 1)
        $grid.Children.Add($assetLabel) | Out-Null
        $grid.Children.Add($assetValue) | Out-Null
        $grid.Children.Add($hostLabel) | Out-Null
        $grid.Children.Add($hostValue) | Out-Null

        $buttonPanel = New-Object System.Windows.Controls.StackPanel -Property @{ Orientation='Horizontal'; HorizontalAlignment='Right' }
        $closeButton = New-Object System.Windows.Controls.Button -Property @{ Content='Close'; MinWidth=86; Padding='14,6'; IsCancel=$true }
        $closeButton.Add_Click({ $window.Close() })
        $buttonPanel.Children.Add($closeButton) | Out-Null
        $panel.Children.Add($buttonPanel) | Out-Null

        $window.Add_ContentRendered({ $closeButton.Focus() | Out-Null })
        $window.ShowDialog() | Out-Null
    }

    function Get-RemoteDeviceSerials {
        param([string]$ComputerName,[Nullable[bool]]$PingSucceeded=$null)
        $result = [pscustomobject]@{ ComputerSerial=$null; MonitorSerials=@(); Offline=$false }
        if ([string]::IsNullOrWhiteSpace($ComputerName)) { return $result }
        $online = if ($null -ne $PingSucceeded) { [bool]$PingSucceeded } else { Test-ComputerPingable -ComputerName $ComputerName }
        if (-not $online) { $result.Offline = $true; return $result }
        try {
            $bios = Get-CimInstance -ClassName Win32_BIOS -ComputerName $ComputerName -ErrorAction Stop
            if ($bios -and $bios.SerialNumber) { $result.ComputerSerial = $bios.SerialNumber.Trim() }
        } catch {}
        if (-not $result.ComputerSerial) {
            try {
                $csprod = Get-CimInstance -ClassName Win32_ComputerSystemProduct -ComputerName $ComputerName -ErrorAction Stop
                if ($csprod -and $csprod.IdentifyingNumber) { $result.ComputerSerial = $csprod.IdentifyingNumber.Trim() }
            } catch {}
        }
        try {
            $monitorData = Get-CimInstance -Namespace root\wmi -ClassName WmiMonitorID -ComputerName $ComputerName -ErrorAction Stop
            foreach ($m in $monitorData) {
                $serial = Convert-WmiIdToString -Id $m.SerialNumberID
                if (-not [string]::IsNullOrWhiteSpace($serial)) { $result.MonitorSerials += $serial.Trim() }
            }
        } catch {}
        return $result
    }

    function Reset-AssociatedSerialValidation {
        param([object[]]$Items)
        foreach ($item in @($Items)) {
            if (-not $item) { continue }
            $item.SerialForeground = '#1F2937'
            $item.SerialToolTip = ''
        }
    }

    function Apply-AssociatedDeviceValidation {
        param([object[]]$Items,[pscustomobject]$WmiData)
        if (-not $WmiData) { return }
        $computerTypes = @('COMPUTER','DESKTOP','LAPTOP','TABLET','THIN CLIENT','TANGENT')
        $monitorSerials = @()
        try { if ($WmiData.MonitorSerials) { $monitorSerials = @($WmiData.MonitorSerials) } } catch {}
        $computerSerial = ''
        try { if ($WmiData.ComputerSerial) { $computerSerial = $WmiData.ComputerSerial.Trim() } } catch {}
        Reset-AssociatedSerialValidation -Items $Items
        foreach ($item in @($Items)) {
            if (-not $item -or [string]::IsNullOrWhiteSpace($item.Serial)) { continue }
            $targets = @()
            $tooltip = ''
            $type = if ($item.Type) { $item.Type.Trim().ToUpper() } else { '' }
            if ($computerTypes -contains $type) {
                if ($computerSerial) { $targets = @($computerSerial.Trim().ToUpper()); $tooltip = "Detected computer serial: $computerSerial" }
                else { $tooltip = 'No computer serial retrieved from WMI.' }
            } elseif ($type -eq 'MONITOR') {
                foreach ($m in $monitorSerials) { if ($m) { $targets += $m.Trim().ToUpper() } }
                if ($targets.Count -gt 0) { $tooltip = 'Detected monitor serials: ' + ($monitorSerials -join ', ') }
                else { $tooltip = 'No monitor serials retrieved from WMI/EDID.' }
            } else { continue }
            $isMatch = $false
            foreach ($target in $targets) { if ($item.Serial.Trim().ToUpper() -eq $target) { $isMatch = $true; break } }
            $item.SerialForeground = if ($isMatch) { '#228B22' } else { '#CD5C5C' }
            $item.SerialToolTip = $tooltip
        }
    }

    function Get-UnlinkedMonitorSerials {
        param([object[]]$Items,[object[]]$MonitorSerials)
        $linked = New-Object System.Collections.ArrayList
        foreach ($item in @($Items)) {
            if (-not $item -or $item.Type -ne 'Monitor' -or [string]::IsNullOrWhiteSpace($item.Serial)) { continue }
            $normalized = $item.Serial.Trim().ToUpper()
            if (-not $linked.Contains($normalized)) { [void]$linked.Add($normalized) }
        }
        $missing = New-Object System.Collections.ArrayList
        foreach ($serial in @($MonitorSerials)) {
            if ([string]::IsNullOrWhiteSpace($serial)) { continue }
            $normalized = $serial.Trim().ToUpper()
            if (-not $linked.Contains($normalized) -and -not $missing.Contains($normalized)) { [void]$missing.Add($normalized) }
        }
        return $missing
    }

    function Validate-AssociatedDevices {
        param([hashtable]$Ui,[pscustomobject]$ParentDevice,[pscustomobject]$Inventory,[string]$ResolvedXamlPath)
        if (-not $ParentDevice -or [string]::IsNullOrWhiteSpace($ParentDevice.Name)) { [System.Windows.MessageBox]::Show('Enter a device name before validating.','Validate Devices') | Out-Null; return }
        $pingable = Test-ComputerPingable -ComputerName $ParentDevice.Name
        if (-not $pingable) { [System.Windows.MessageBox]::Show('Device is not pingable.','Validate Devices') | Out-Null; return }
        $wmiData = Get-RemoteDeviceSerials -ComputerName $ParentDevice.Name -PingSucceeded $pingable
        if ($wmiData.Offline) { [System.Windows.MessageBox]::Show('Device is not pingable.','Validate Devices') | Out-Null; return }
        $items = @($Ui.AssociatedDevicesDataGrid.ItemsSource)
        Apply-AssociatedDeviceValidation -Items $items -WmiData $wmiData
        $Ui.AssociatedDevicesDataGrid.Items.Refresh()
        $missingMonitors = @(Get-UnlinkedMonitorSerials -Items $items -MonitorSerials @($wmiData.MonitorSerials))
        if ($missingMonitors.Count -gt 0) {
            $serialToLink = $missingMonitors | Select-Object -First 1
            $changed = Show-AddPeripheralDialog -Ui $Ui -ParentDevice $ParentDevice -Inventory $Inventory -ResolvedXamlPath $ResolvedXamlPath -DefaultSearchText $serialToLink -InfoMessage 'This monitor is connected but not linked. Click Add to link.'
            if ($changed) { $Ui.AssociatedDevicesDataGrid.ItemsSource = Build-AssociatedDevices -Device $ParentDevice -Inventory $Inventory }
        }
    }

    function Show-AddPeripheralDialog {
        param([hashtable]$Ui,[pscustomobject]$ParentDevice,[pscustomobject]$Inventory,[string]$ResolvedXamlPath,[string]$DefaultSearchText='',[string]$InfoMessage='')
        if (-not $ParentDevice) { return $false }
        $window = New-Object System.Windows.Window
        $window.Title = 'Add Peripheral (Name/Asset/Serial)'
        $window.SizeToContent = 'WidthAndHeight'
        $window.WindowStartupLocation = 'CenterOwner'
        $window.ResizeMode = 'NoResize'
        $window.Owner = [System.Windows.Window]::GetWindow($Ui.MainTabControl)

        $panel = New-Object System.Windows.Controls.StackPanel
        $panel.Margin = '12'
        $window.Content = $panel
        $panel.Children.Add((New-Object System.Windows.Controls.TextBlock -Property @{ Text='Universal Search (name / asset / serial):'; Margin='0,0,0,4' })) | Out-Null
        if (-not [string]::IsNullOrWhiteSpace($InfoMessage)) { $panel.Children.Add((New-Object System.Windows.Controls.TextBlock -Property @{ Text=$InfoMessage; Foreground='#B45309'; TextWrapping='Wrap'; Margin='0,0,0,8' })) | Out-Null }
        $txt = New-Object System.Windows.Controls.TextBox
        $txt.MinWidth = 320
        $txt.Text = $DefaultSearchText
        $txt.Margin = '0,0,0,6'
        $panel.Children.Add($txt) | Out-Null
        $hint = New-Object System.Windows.Controls.TextBlock -Property @{ Text='Press Enter to search by cart name, asset tag, or serial number.'; Foreground='#64748B'; Margin='0,0,0,8' }
        $panel.Children.Add($hint) | Out-Null
        $previewGroup = New-Object System.Windows.Controls.GroupBox -Property @{ Header='Peripheral Preview'; Margin='0,0,0,10'; Padding='10' }
        $previewGrid = New-Object System.Windows.Controls.Grid
        $previewGrid.ColumnDefinitions.Add((New-Object System.Windows.Controls.ColumnDefinition -Property @{ Width='Auto' })) | Out-Null
        $previewGrid.ColumnDefinitions.Add((New-Object System.Windows.Controls.ColumnDefinition -Property @{ Width='*' })) | Out-Null
        $previewGroup.Content = $previewGrid
        $panel.Children.Add($previewGroup) | Out-Null

        $previewValues = @{}
        $addPreviewRow = {
            param([int]$Row,[string]$Label,[string]$Key)
            $previewGrid.RowDefinitions.Add((New-Object System.Windows.Controls.RowDefinition -Property @{ Height='Auto' })) | Out-Null
            $lbl = New-Object System.Windows.Controls.TextBlock -Property @{ Text=$Label; Margin='0,0,28,4' }
            [System.Windows.Controls.Grid]::SetRow($lbl,$Row); [System.Windows.Controls.Grid]::SetColumn($lbl,0)
            $previewGrid.Children.Add($lbl) | Out-Null
            $val = New-Object System.Windows.Controls.TextBlock -Property @{ Text=''; Margin='0,0,0,4' }
            [System.Windows.Controls.Grid]::SetRow($val,$Row); [System.Windows.Controls.Grid]::SetColumn($val,1)
            $previewGrid.Children.Add($val) | Out-Null
            $previewValues[$Key] = $val
        }
        & $addPreviewRow 0 'Type:' 'Type'
        & $addPreviewRow 1 'Name:' 'Name'
        & $addPreviewRow 2 'Parent:' 'Parent'
        & $addPreviewRow 3 'Asset Tag:' 'AssetTag'
        & $addPreviewRow 4 'Serial Number:' 'Serial'
        & $addPreviewRow 5 'RITM:' 'RITM'
        & $addPreviewRow 6 'Retire:' 'Retire'
        $buttons = New-Object System.Windows.Controls.StackPanel -Property @{ Orientation='Horizontal'; HorizontalAlignment='Right' }
        $addButton = New-Object System.Windows.Controls.Button -Property @{ Content='Add'; IsEnabled=$false; Margin='0,0,8,0'; MinWidth=90 }
        $cancelButton = New-Object System.Windows.Controls.Button -Property @{ Content='Cancel'; MinWidth=90 }
        $buttons.Children.Add($addButton) | Out-Null
        $buttons.Children.Add($cancelButton) | Out-Null
        $panel.Children.Add($buttons) | Out-Null

        $lookupResult = $null
        $updatePreview = {
            param($result)
            $lookupResult = $result
            foreach ($value in $previewValues.Values) { $value.Text = '' }
            if (-not $result -or -not $result.Candidate) {
                $previewValues.Type.Text = if ($result -and $result.NormalizedInput) { '(not found)' } else { '' }
                $addButton.IsEnabled = $false
                return
            }
            $c = $result.Candidate
            $proposedName = Get-ProposedPeripheralName -Candidate $c -ParentDevice $ParentDevice
            if ([string]::IsNullOrWhiteSpace($proposedName)) { $proposedName = $c.Name }
            $proposedParent = $ParentDevice.AssetTag
            $currentParentDisplay = Get-ParentPreviewDisplay -Token $c.Parent -FallbackParent $ParentDevice -Inventory $Inventory
            $proposedParentDisplay = Get-ParentPreviewDisplay -Token $proposedParent -FallbackParent $ParentDevice -Inventory $Inventory
            $previewValues.Type.Text = $c.DetectedType
            $previewValues.Name.Text = if ($c.Name -ne $proposedName) { "$($c.Name)     ---->     $proposedName" } else { $c.Name }
            $previewValues.Parent.Text = if ($currentParentDisplay -ne $proposedParentDisplay) { "$currentParentDisplay     ---->     $proposedParentDisplay" } else { $currentParentDisplay }
            $previewValues.AssetTag.Text = $c.AssetTag
            $previewValues.Serial.Text = $c.Serial
            $previewValues.RITM.Text = $c.RITM
            $previewValues.Retire.Text = $c.RetireDate
            $addButton.IsEnabled = $true
        }

        $txt.Add_TextChanged({ $lookupResult = $null; foreach ($value in $previewValues.Values) { $value.Text = '' }; $addButton.IsEnabled = $false })
        $txt.Add_KeyDown({
            if ($_.Key -eq [System.Windows.Input.Key]::Enter -or $_.Key -eq [System.Windows.Input.Key]::Return) {
                $_.Handled = $true
                $result = Resolve-AssociatedPeripheralLookup -Query $txt.Text -Inventory $Inventory
                & $updatePreview $result
            }
        })
        $cancelButton.Add_Click({ $window.DialogResult = $false; $window.Close() })
        $addButton.Add_Click({
            $result = $lookupResult
            if (-not $result) { $result = Resolve-AssociatedPeripheralLookup -Query $txt.Text -Inventory $Inventory; & $updatePreview $result }
            if (-not $result -or -not $result.Candidate) { return }
            $target = $result.Candidate
            $oldParent = $target.Parent
            $oldName = $target.Name
            $newName = Get-ProposedPeripheralName -Candidate $target -ParentDevice $ParentDevice
            if ([string]::IsNullOrWhiteSpace($newName)) { $newName = $target.Name }
            foreach ($collectionName in @('Monitors','Carts','Mics','Scanners')) {
                foreach ($row in $Inventory.$collectionName) {
                    $at = Get-FieldValue -Row $row -Names @('asset_tag')
                    if ($at -and $target.AssetTag -and $at.Trim().ToUpper() -eq $target.AssetTag.Trim().ToUpper()) {
                        Set-RowFieldValue -Row $row -Name 'u_parent_asset' -Value $ParentDevice.AssetTag
                        Set-RowFieldValue -Row $row -Name 'name' -Value $newName
                    }
                }
            }
            Add-CmdbAssociationUpdate -ResolvedXamlPath $ResolvedXamlPath -Candidate $target -OldParent $oldParent -NewParent $ParentDevice.AssetTag -OldName $oldName -NewName $newName -Action 'Link' -Inventory $Inventory
            Build-InventoryIndices -Inventory $Inventory
            $window.DialogResult = $true
            $window.Close()
        })
        if (-not [string]::IsNullOrWhiteSpace($DefaultSearchText)) { $result = Resolve-AssociatedPeripheralLookup -Query $DefaultSearchText -Inventory $Inventory; & $updatePreview $result }
        $null = $window.ShowDialog()
        return [bool]$window.DialogResult
    }

    function Get-CmdbLink {
        param([string]$DeviceType,[string]$AssetTag)
        if ([string]::IsNullOrWhiteSpace($AssetTag)) { return '' }

        $assetValue = $AssetTag.Trim()
        $innerPath = $null
        $peripheralTypes = @('Monitor','Mic','Scanner','Microphone')
        $computerTypes = @('Computer','Tangent','Desktop','Laptop','Thin Client')

        if ($peripheralTypes -contains $DeviceType) {
            $innerPath = 'cmdb_ci_peripheral_list.do?sysparm_first_row=1&sysparm_query=GOTOasset_tagLIKE{0}&sysparm_query_encoded=GOTOasset_tagLIKE{0}&sysparm_view='
        } elseif ($computerTypes -contains $DeviceType) {
            $innerPath = 'cmdb_ci_computer_list.do?sysparm_first_row=1&sysparm_query=companyINjavascript:new inccompanysearchChange().getFilter();^GOTOasset_tagLIKE{0}&sysparm_query_encoded=companyINjavascript:new inccompanysearchChange().getFilter();^GOTOasset_tagLIKE{0}&sysparm_view='
        } elseif ($DeviceType -eq 'Cart') {
            $innerPath = 'u_cmdb_ci_mobile_carts_list.do?sysparm_first_row=1&sysparm_query=companyINjavascript:new inccompanysearchChange().getFilter();^operational_status!=6^GOTOasset_tagLIKE{0}&sysparm_query_encoded=companyINjavascript:new inccompanysearchChange().getFilter();^operational_status!=6^GOTOasset_tagLIKE{0}&sysparm_view='
        }

        if (-not $innerPath) { return '' }
        $expandedInnerPath = [string]::Format($innerPath,$assetValue)
        $encodedInnerPath = [System.Uri]::EscapeDataString($expandedInnerPath)
        return "https://healthbc.service-now.com/nav_to.do?uri=$encodedInnerPath"
    }

    function Build-QueryData {
        param([pscustomobject]$Device,[pscustomobject]$Inventory)
        return [pscustomobject]@{ Device=$Device; Associated=(Build-AssociatedDevices -Device $Device -Inventory $Inventory); Nearby=(Build-NearbyDevices -Device $Device -Inventory $Inventory) }
    }



    function Get-ExpectedDeviceName {
        param([pscustomobject]$Device,[pscustomobject]$ParentDevice)
        if (-not $Device -or -not $ParentDevice) { return $null }
        switch ($Device.DetectedType) {
            'Monitor' { return $ParentDevice.Name }
            'Mic' { return "$($ParentDevice.Name)-Mic" }
            'Scanner' { return "$($ParentDevice.Name)-SCN" }
            'Cart' { return "$($ParentDevice.Name)-CRT" }
            default { return $null }
        }
    }

    function Test-DeviceNameNeedsFix {
        param([pscustomobject]$Device,[pscustomobject]$ParentDevice)
        $expected = Get-ExpectedDeviceName -Device $Device -ParentDevice $ParentDevice
        if ([string]::IsNullOrWhiteSpace($expected) -or [string]::IsNullOrWhiteSpace($Device.Name)) { return $false }
        return ($Device.Name.Trim().ToUpper() -ne $expected.Trim().ToUpper())
    }

    function Update-FixNameButtonState {
        param([hashtable]$Ui,[pscustomobject]$Device,[pscustomobject]$ParentDevice)
        if (-not $Ui.FixNameButton) { return }
        if (Test-DeviceNameNeedsFix -Device $Device -ParentDevice $ParentDevice) {
            $Ui.FixNameButton.Visibility = 'Visible'
            $Ui.FixNameButton.IsEnabled = $true
        }
        else {
            $Ui.FixNameButton.Visibility = 'Collapsed'
            $Ui.FixNameButton.IsEnabled = $false
        }
    }

    function Update-InventoryDeviceName {
        param([pscustomobject]$Inventory,[pscustomobject]$Device,[string]$NewName)
        if (-not $Inventory -or -not $Device -or [string]::IsNullOrWhiteSpace($Device.AssetTag)) { return }
        foreach ($collectionName in @('Monitors','Carts','Mics','Scanners')) {
            foreach ($row in $Inventory.$collectionName) {
                $assetTag = Get-FieldValue -Row $row -Names @('asset_tag','AssetTag')
                if ($assetTag -and $assetTag.Trim().ToUpper() -eq $Device.AssetTag.Trim().ToUpper()) {
                    Set-RowFieldValue -Row $row -Name 'name' -Value $NewName
                    return
                }
            }
        }
    }

    function Remove-SelectedPeripheralAssociations {
        param([hashtable]$Ui,[pscustomobject]$Inventory,[string]$ResolvedXamlPath)
        if (-not $Ui.AssociatedDevicesDataGrid -or -not $Inventory) { return $false }
        $selectedRows = @($Ui.AssociatedDevicesDataGrid.SelectedItems)
        if ($selectedRows.Count -eq 0) { return $false }

        $changed = $false
        foreach ($selected in $selectedRows) {
            if (-not $selected) { continue }
            $role = [string]$selected.Role
            if (($role -ne 'Child') -and ($role -ne 'Grandchild')) { continue }
            if ([string]::IsNullOrWhiteSpace($selected.AssetTag)) { continue }

            $found = $false
            foreach ($collectionName in @('Monitors','Carts','Mics','Scanners')) {
                foreach ($row in $Inventory.$collectionName) {
                    $assetTag = Get-FieldValue -Row $row -Names @('asset_tag','AssetTag')
                    if ([string]::IsNullOrWhiteSpace($assetTag) -or $assetTag.Trim().ToUpper() -ne $selected.AssetTag.Trim().ToUpper()) { continue }

                    $deviceType = if ($collectionName -eq 'Carts') { 'Cart' } elseif ($collectionName -eq 'Mics') { 'Mic' } elseif ($collectionName -eq 'Scanners') { 'Scanner' } else { 'Monitor' }
                    $device = ConvertTo-DeviceRecord -Row $row -DetectedType $deviceType
                    $oldParent = $device.Parent
                    $oldName = $device.Name
                    $newName = if ([string]::IsNullOrWhiteSpace($device.Serial)) { $device.Name } else { $device.Serial }

                    Set-RowFieldValue -Row $row -Name 'u_parent_asset' -Value $null
                    Set-RowFieldValue -Row $row -Name 'name' -Value $newName
                    $device.Parent = $null
                    $device.Name = $newName

                    Add-CmdbAssociationUpdate -ResolvedXamlPath $ResolvedXamlPath -Candidate $device -OldParent $oldParent -NewParent $null -OldName $oldName -NewName $newName -Action 'Unlink' -Inventory $Inventory
                    $changed = $true
                    $found = $true
                    break
                }
                if ($found) { break }
            }
        }

        if ($changed) { Build-InventoryIndices -Inventory $Inventory }
        return $changed
    }


    function Normalize-LocationValue {
        param([string]$Value)
        if ([string]::IsNullOrWhiteSpace($Value)) { return '' }
        return (($Value -replace ([char]0xFEFF), '').Trim().ToUpperInvariant())
    }

    function Sort-LocationFloors {
        param([object[]]$Floors)
        return @($Floors | Sort-Object { if ($_ -match '^-?\d+$') { '{0:D8}' -f [int]$_ } else { [string]$_ } } -Unique)
    }

    function Get-LocationFieldValue {
        param([object]$Row,[string]$Name)
        return Get-FieldValue -Row $Row -Names @($Name,(([char]0xFEFF) + $Name))
    }

    function Get-LocationRows {
        param([pscustomobject]$Inventory)
        $rows = New-Object System.Collections.Generic.List[object]
        $seen = @{}
        $add = {
            param($city,$location,$building,$floor,$room,$department)
            $key = '{0}|{1}|{2}|{3}|{4}|{5}' -f (Normalize-LocationValue $city),(Normalize-LocationValue $location),(Normalize-LocationValue $building),(Normalize-LocationValue $floor),(Normalize-LocationValue $room),(Normalize-LocationValue $department)
            if (-not $seen.ContainsKey($key)) {
                $seen[$key] = $true
                [void]$rows.Add([pscustomobject]@{ City=[string]$city; Location=[string]$location; Building=[string]$building; Floor=[string]$floor; Room=[string]$room; Department=[string]$department })
            }
        }
        foreach ($row in @($Inventory.Locations)) {
            & $add (Get-LocationFieldValue $row 'City') (Get-LocationFieldValue $row 'Location') (Get-LocationFieldValue $row 'Building') (Get-LocationFieldValue $row 'Floor') (Get-LocationFieldValue $row 'Room') (Get-LocationFieldValue $row 'Department')
        }
        foreach ($row in @($Inventory.Computers)) {
            & $add (Get-FieldValue $row @('location.city','City')) (Get-FieldValue $row @('location','Location')) (Get-FieldValue $row @('u_building','Building')) (Get-FieldValue $row @('u_floor','Floor')) (Get-FieldValue $row @('u_room','Room')) (Get-FieldValue $row @('u_department_location','Department'))
        }
        return @($rows)
    }

    function Test-LocationColumnValue {
        param([pscustomobject]$Inventory,[string]$Value,[string]$Column)
        if ([string]::IsNullOrWhiteSpace($Value)) { return $false }
        $n = Normalize-LocationValue $Value
        foreach ($row in @($Inventory.Locations)) {
            if ((Normalize-LocationValue (Get-LocationFieldValue $row $Column)) -eq $n) { return $true }
        }
        return $false
    }

    function Extract-RoomCode {
        param([string]$Value)
        $n = Normalize-LocationValue $Value
        $m = [regex]::Match($n, '^[A-Z0-9]+')
        if ($m.Success) { return $m.Value }
        return ''
    }

    function Test-LocationRoomValue {
        param([pscustomobject]$Inventory,[string]$Value)
        if ([string]::IsNullOrWhiteSpace($Value)) { return $false }
        $n = Normalize-LocationValue $Value
        $code = Extract-RoomCode $Value
        foreach ($row in @($Inventory.Locations)) {
            $room = Get-LocationFieldValue $row 'Room'
            if ((Normalize-LocationValue $room) -eq $n) { return $true }
            if ($code -and (Extract-RoomCode $room) -eq $code) { return $true }
        }
        return $false
    }

    function Set-LocationValidationStyle {
        param([hashtable]$Ui,[pscustomobject]$Inventory)
        $validBrush = New-Brush '#CCF2D3'; $validBorder = New-Brush '#7CE0A6'
        $badBrush = New-Brush '#FCE3E5'; $badBorder = New-Brush '#F5A3AA'
        $checks = @(
            @($Ui.CityTextBox,       (Test-LocationColumnValue $Inventory $Ui.CityTextBox.Text 'City')),
            @($Ui.LocationTextBox,   (Test-LocationColumnValue $Inventory $Ui.LocationTextBox.Text 'Location')),
            @($Ui.BuildingTextBox,   (Test-LocationColumnValue $Inventory $Ui.BuildingTextBox.Text 'Building')),
            @($Ui.FloorTextBox,      (Test-LocationColumnValue $Inventory $Ui.FloorTextBox.Text 'Floor')),
            @($Ui.RoomTextBox,       (Test-LocationRoomValue $Inventory $Ui.RoomTextBox.Text)),
            @($Ui.DepartmentTextBox, (Test-LocationColumnValue $Inventory $Ui.DepartmentTextBox.Text 'Department'))
        )
        foreach ($item in $checks) {
            $box = $item[0]; $ok = [bool]$item[1]
            $box.Background = if ($ok) { $validBrush } else { $badBrush }
            $box.BorderBrush = if ($ok) { $validBorder } else { $badBorder }
        }
    }

    function Get-ValidLocationSelection {
        param([string]$Value,[object[]]$Items)
        $n = Normalize-LocationValue $Value
        if (-not $n) { return $null }
        foreach ($item in @($Items)) { if ((Normalize-LocationValue $item) -eq $n) { return [string]$item } }
        return $null
    }

    function Set-ComboItems {
        param([System.Windows.Controls.ComboBox]$Combo,[object[]]$Items,[string]$Text)
        if (-not $Combo) { return }
        $targetText = if ($null -ne $Text) { [string]$Text } else { '' }
        $Combo.Items.Clear()
        foreach ($item in @($Items)) {
            if ($null -eq $item) { continue }
            $itemText = [string]$item
            if ([string]::IsNullOrWhiteSpace($itemText)) { continue }
            [void]$Combo.Items.Add($itemText)
        }
        if ($Combo.Text -ne $targetText) { $Combo.Text = $targetText }
        if ([string]::IsNullOrWhiteSpace($targetText)) {
            $Combo.SelectedIndex = -1
            return
        }
        try {
            $matchIndex = -1
            for ($i = 0; $i -lt $Combo.Items.Count; $i++) {
                if ((Normalize-LocationValue $Combo.Items[$i]) -eq (Normalize-LocationValue $targetText)) {
                    $matchIndex = $i
                    break
                }
            }
            if ($matchIndex -ge 0) { $Combo.SelectedIndex = $matchIndex }
        } catch {}
    }

    function Filter-LocationRows {
        param([object[]]$Rows,[string]$City,[string]$Location,[string]$Building,[string]$Floor)
        $filtered = @($Rows)
        if (-not [string]::IsNullOrWhiteSpace($City)) {
            $nCity = Normalize-LocationValue $City
            $filtered = @($filtered | Where-Object { (Normalize-LocationValue $_.City) -eq $nCity })
        }
        if (-not [string]::IsNullOrWhiteSpace($Location)) {
            $nLocation = Normalize-LocationValue $Location
            $filtered = @($filtered | Where-Object { (Normalize-LocationValue $_.Location) -eq $nLocation })
        }
        if (-not [string]::IsNullOrWhiteSpace($Building)) {
            $nBuilding = Normalize-LocationValue $Building
            $filtered = @($filtered | Where-Object { (Normalize-LocationValue $_.Building) -eq $nBuilding })
        }
        if (-not [string]::IsNullOrWhiteSpace($Floor)) {
            $nFloor = Normalize-LocationValue $Floor
            $filtered = @($filtered | Where-Object { (Normalize-LocationValue $_.Floor) -eq $nFloor })
        }
        return @($filtered)
    }

    function Populate-LocationCombos {
        param([hashtable]$Ui,[pscustomobject]$Inventory,[string]$ChangedLevel='')
        if (-not $Ui -or -not $Inventory) { return }
        $script:IsPopulatingLocationCombos = $true
        try {
            $rows = @(Get-LocationRows -Inventory $Inventory)

            $city = [string]$Ui.CityComboBox.Text
            $loc = [string]$Ui.LocationComboBox.Text
            $building = [string]$Ui.BuildingComboBox.Text
            $floor = [string]$Ui.FloorComboBox.Text
            $room = [string]$Ui.RoomComboBox.Text
            $department = [string]$Ui.DepartmentComboBox.Text

            if ($ChangedLevel -eq 'City') { $loc = ''; $building = ''; $floor = ''; $room = '' }
            elseif ($ChangedLevel -eq 'Location') { $building = ''; $floor = ''; $room = '' }
            elseif ($ChangedLevel -eq 'Building') { $floor = ''; $room = '' }
            elseif ($ChangedLevel -eq 'Floor') { $room = '' }

            $cities = @($rows | ForEach-Object { $_.City } | Where-Object { -not [string]::IsNullOrWhiteSpace([string]$_) } | Sort-Object -Unique)
            Set-ComboItems $Ui.CityComboBox $cities $city
            $validCity = Get-ValidLocationSelection $Ui.CityComboBox.Text $cities

            # Match the old tool's behavior: only filter a lower level when the higher-level
            # text is a valid LocationMaster value. This keeps dropdowns populated even when
            # a device has a legacy/free-form city, location, building, floor, or room value.
            $locRows = Filter-LocationRows -Rows $rows -City $validCity
            $locs = @($locRows | ForEach-Object { $_.Location } | Where-Object { -not [string]::IsNullOrWhiteSpace([string]$_) } | Sort-Object -Unique)
            Set-ComboItems $Ui.LocationComboBox $locs $loc
            $validLoc = Get-ValidLocationSelection $Ui.LocationComboBox.Text $locs

            $bldRows = if ($validLoc) { Filter-LocationRows -Rows $locRows -Location $validLoc } else { @($locRows) }
            $buildings = @($bldRows | ForEach-Object { $_.Building } | Where-Object { -not [string]::IsNullOrWhiteSpace([string]$_) } | Sort-Object -Unique)
            Set-ComboItems $Ui.BuildingComboBox $buildings $building
            $validBuilding = Get-ValidLocationSelection $Ui.BuildingComboBox.Text $buildings

            $floorRows = if ($validBuilding) { Filter-LocationRows -Rows $bldRows -Building $validBuilding } else { @($bldRows) }
            $floors = Sort-LocationFloors @($floorRows | ForEach-Object { $_.Floor } | Where-Object { -not [string]::IsNullOrWhiteSpace([string]$_) })
            Set-ComboItems $Ui.FloorComboBox $floors $floor
            $validFloor = Get-ValidLocationSelection $Ui.FloorComboBox.Text $floors

            $roomRows = if ($validFloor) { Filter-LocationRows -Rows $floorRows -Floor $validFloor } else { @($floorRows) }
            $rooms = @($roomRows | ForEach-Object { $_.Room } | Where-Object { -not [string]::IsNullOrWhiteSpace([string]$_) } | Sort-Object -Unique)
            Set-ComboItems $Ui.RoomComboBox $rooms $room

            $departmentRows = if (-not [string]::IsNullOrWhiteSpace($Ui.RoomComboBox.Text)) {
                $validRoom = Get-ValidLocationSelection $Ui.RoomComboBox.Text $rooms
                if ($validRoom) { @($roomRows | Where-Object { (Normalize-LocationValue $_.Room) -eq (Normalize-LocationValue $validRoom) }) } else { @($roomRows) }
            } else { @($roomRows) }
            $departments = @($departmentRows | ForEach-Object { $_.Department } | Where-Object { -not [string]::IsNullOrWhiteSpace([string]$_) } | Sort-Object -Unique)
            if ($departments.Count -eq 0) { $departments = @($rows | ForEach-Object { $_.Department } | Where-Object { -not [string]::IsNullOrWhiteSpace([string]$_) } | Sort-Object -Unique) }
            Set-ComboItems $Ui.DepartmentComboBox $departments $department
        } finally { $script:IsPopulatingLocationCombos = $false }
    }

    function Set-SelectedSummaryDevice {
        param([hashtable]$Ui,[pscustomobject]$Device,[pscustomobject]$Inventory)
        if (-not $Device) { return }
        $parentDevice = Resolve-ParentDevice -Device $Device -Inventory $Inventory
        $locationDevice = if ($parentDevice) { $parentDevice } else { $Device }
        $Ui.SelectedDeviceText.Text = $Device.Name
        Set-DisplayText -Ui $Ui -BaseName 'DetectedType' -Value $Device.DetectedType
        Set-DisplayText -Ui $Ui -BaseName 'HostName' -Value $Device.Name
        Set-DisplayText -Ui $Ui -BaseName 'AssetTag' -Value $Device.AssetTag
        Set-DisplayText -Ui $Ui -BaseName 'Serial' -Value $Device.Serial
        Set-DisplayText -Ui $Ui -BaseName 'Parent' -Value $Device.Parent
        Set-DisplayText -Ui $Ui -BaseName 'Ritm' -Value $Device.RITM
        Set-DisplayText -Ui $Ui -BaseName 'Retire' -Value (Format-DateLong $Device.RetireDate)
        Set-LastRoundedDisplay -Ui $Ui -LastRoundedRaw $locationDevice.LastRounded
        $Ui.CityTextBox.Text = $locationDevice.City
        $Ui.LocationTextBox.Text = $locationDevice.Location
        $Ui.BuildingTextBox.Text = $locationDevice.Building
        $Ui.FloorTextBox.Text = $locationDevice.Floor
        $Ui.RoomTextBox.Text = $locationDevice.Room
        $Ui.DepartmentTextBox.Text = $locationDevice.Department
        Set-LocationValidationStyle -Ui $Ui -Inventory $Inventory
        $script:AppState.SelectedSummaryDevice = $Device
        $script:AppState.SelectedSummaryParent = $parentDevice
        Update-FixNameButtonState -Ui $Ui -Device $Device -ParentDevice $parentDevice
    }

    function Set-PrimaryDeviceBindings {
        param([hashtable]$Ui,[pscustomobject]$Device,[pscustomobject]$Inventory)
        Set-SelectedSummaryDevice -Ui $Ui -Device $Device -Inventory $Inventory
    }

    function Start-QueryDataPopulationAsync {
        param([hashtable]$Ui,[pscustomobject]$Device,[pscustomobject]$Inventory,[string]$QueryToken)
        [System.Threading.Tasks.Task]::Run([Action]{
            $associated = Build-AssociatedDevices -Device $Device -Inventory $Inventory
            $Ui.MainTabControl.Dispatcher.BeginInvoke([Action]{
                if ($script:AppState.CurrentQueryToken -ne $QueryToken) { return }
                $Ui.AssociatedDevicesDataGrid.ItemsSource = $associated
                $Ui.NearbyDataGrid.ItemsSource = @()
                Set-ControlText -Control $Ui.NearbyScopeSummaryText -Value 'Nearby disabled'
                $script:AppState.SampleData = [pscustomobject]@{ Device=$Device; Associated=$associated; Nearby=@() }
            }) | Out-Null
        }) | Out-Null
    }

    function New-SampleData {
        # fallback sample data when CSV sources are unavailable
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
            [pscustomobject]@{ Role='Parent'; Type='Tangent'; Name='AO400568'; AssetTag='HSS-8093577'; Serial='C24102M031'; SerialForeground='#1F2937'; SerialToolTip=''; RITM='TRP - 26 May 2025'; Retire='31 May 2028' },
            [pscustomobject]@{ Role='Child'; Type='Cart'; Name='AO400568-CRT'; AssetTag='CO09167'; Serial='1896875-0016'; SerialForeground='#1F2937'; SerialToolTip=''; RITM='-'; Retire='-' }
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
        $Ui.SelectedDeviceText.Text = $device.Name
        Set-DisplayText -Ui $Ui -BaseName 'DetectedType' -Value $device.DetectedType
        Set-DisplayText -Ui $Ui -BaseName 'HostName' -Value $device.Name
        Set-DisplayText -Ui $Ui -BaseName 'AssetTag' -Value $device.AssetTag
        Set-DisplayText -Ui $Ui -BaseName 'Serial' -Value $device.Serial
        Set-DisplayText -Ui $Ui -BaseName 'Parent' -Value $device.Parent
        Set-DisplayText -Ui $Ui -BaseName 'Ritm' -Value $device.RITM
        Set-DisplayText -Ui $Ui -BaseName 'Retire' -Value (Format-DateLong $device.RetireDate)
        Set-LastRoundedDisplay -Ui $Ui -LastRoundedRaw $device.LastRounded
        $Ui.CityTextBox.Text = $device.City
        $Ui.LocationTextBox.Text = $device.Location
        $Ui.BuildingTextBox.Text = $device.Building
        $Ui.FloorTextBox.Text = $device.Floor
        $Ui.RoomTextBox.Text = $device.Room
        $Ui.DepartmentTextBox.Text = $device.Department
        if ($script:AppState -and $script:AppState.Inventory) { Set-LocationValidationStyle -Ui $Ui -Inventory $script:AppState.Inventory }
        $Ui.LastQueryBadgeText.Text = "Queried $(Get-Date -Format 'HH:mm')"
        Set-ControlText -Control $Ui.NearbyScopeSummaryText -Value 'Nearby disabled'
        $Ui.AssociatedDevicesDataGrid.ItemsSource = $SampleData.Associated
        $Ui.NearbyDataGrid.ItemsSource = $SampleData.Nearby
    }

    
    function Update-OnlineStatus {
        param([hashtable]$Ui,[string]$HostName)
        $isOnline = $false
        $latencyMs = $null
        try {
            $result = Test-Connection -ComputerName $HostName -Count 1 -ErrorAction Stop
            if ($result) {
                $isOnline = $true
                $latencyMs = [int][Math]::Round(($result | Select-Object -First 1).ResponseTime)
            }
        } catch {}

        if ($isOnline) {
            $Ui.DeviceOnlineText.Text = 'Online'
            $Ui.DeviceOnlineText.Foreground = New-Brush '#16A34A'
            $Ui.DeviceOnlineDot.Fill = New-Brush '#16A34A'
            $Ui.DeviceResponseTimeText.Text = "($latencyMs ms)"
        }
        else {
            $Ui.DeviceOnlineText.Text = 'Offline'
            $Ui.DeviceOnlineText.Foreground = New-Brush '#BE123C'
            $Ui.DeviceOnlineDot.Fill = New-Brush '#BE123C'
            $Ui.DeviceResponseTimeText.Text = '(No response)'
        }

        $Ui.LastQueryBadgeText.Text = "Queried $(Get-Date -Format 'HH:mm:ss')"
    }

    function Set-OnlineStatusUi {
        param([hashtable]$Ui,[bool]$IsOnline,[Nullable[int]]$LatencyMs)
        if ($IsOnline) {
            $Ui.DeviceOnlineText.Text = 'Online'
            $Ui.DeviceOnlineText.Foreground = New-Brush '#16A34A'
            $Ui.DeviceOnlineDot.Fill = New-Brush '#16A34A'
            $Ui.DeviceResponseTimeText.Text = "($LatencyMs ms)"
        }
        else {
            $Ui.DeviceOnlineText.Text = 'Offline'
            $Ui.DeviceOnlineText.Foreground = New-Brush '#BE123C'
            $Ui.DeviceOnlineDot.Fill = New-Brush '#BE123C'
            $Ui.DeviceResponseTimeText.Text = '(No response)'
        }
        $Ui.LastQueryBadgeText.Text = "Queried $(Get-Date -Format 'HH:mm:ss')"
    }

    function Start-OnlineStatusUpdateAsync {
        param([hashtable]$Ui,[string]$HostName)
        if ([string]::IsNullOrWhiteSpace($HostName)) {
            Set-OnlineStatusUi -Ui $Ui -IsOnline:$false -LatencyMs $null
            return
        }
        [System.Threading.Tasks.Task]::Run([Action]{
            $isOnline = $false
            $latencyMs = $null
            try {
                $result = Test-Connection -ComputerName $HostName -Count 1 -ErrorAction Stop
                if ($result) {
                    $isOnline = $true
                    $latencyMs = [int][Math]::Round(($result | Select-Object -First 1).ResponseTime)
                }
            } catch {}
            $Ui.MainTabControl.Dispatcher.BeginInvoke([Action]{
                Set-OnlineStatusUi -Ui $Ui -IsOnline:$isOnline -LatencyMs $latencyMs
            }) | Out-Null
        }) | Out-Null
    }

    function Increment-Fonts {
        param([System.Windows.DependencyObject]$Root)
        if ($null -eq $Root) { return }
        if ($Root -is [System.Windows.Controls.Control] -or $Root -is [System.Windows.Controls.TextBlock]) {
            $skip = ($Root -is [System.Windows.Controls.TextBlock]) -and ($Root.Text -in @('System','Nearby'))
            if (-not $skip) { $Root.FontSize = [Math]::Max(1, $Root.FontSize + 1) }
        }
        $count = [System.Windows.Media.VisualTreeHelper]::GetChildrenCount($Root)
        for ($i=0; $i -lt $count; $i++) {
            Increment-Fonts -Root ([System.Windows.Media.VisualTreeHelper]::GetChild($Root, $i))
        }
    }

    function Clear-WindowData {
        param([hashtable]$Ui)
        foreach ($name in @('DetectedType','HostName','AssetTag','Serial','Parent','Ritm','Retire')) { Set-DisplayText -Ui $Ui -BaseName $name -Value '' }
        Set-ControlText -Control $Ui.SelectedDeviceText -Value ''
        $Ui.LastRoundedLabelText.Foreground = New-Brush '#64748B'
        Set-ControlText -Control $Ui.LastRoundedText -Value ''
        $Ui.LastRoundedText.Foreground = New-Brush '#64748B'
        $Ui.LastRoundedContainer.Background = New-Brush '#F8FAFC'
        $Ui.LastRoundedContainer.BorderBrush = New-Brush '#D9E1EA'
        $Ui.LastRoundedAttentionBadge.Visibility = 'Collapsed'
        foreach ($box in @($Ui.CityTextBox,$Ui.LocationTextBox,$Ui.BuildingTextBox,$Ui.FloorTextBox,$Ui.RoomTextBox,$Ui.DepartmentTextBox)) { Set-ControlText -Control $box -Value ''; $box.Background = New-Brush '#CCF2D3'; $box.BorderBrush = New-Brush '#7CE0A6' }
        Set-ControlText -Control $Ui.NearbyScopeSummaryText -Value 'Nearby disabled'
        $Ui.AssociatedDevicesDataGrid.ItemsSource = @()
        $Ui.NearbyDataGrid.ItemsSource = @()
        Set-ControlText -Control $Ui.DeviceOnlineText -Value 'Ready'
        $Ui.DeviceOnlineText.Foreground = New-Brush '#64748B'
        $Ui.DeviceOnlineDot.Fill = New-Brush '#94A3B8'
        Set-ControlText -Control $Ui.DeviceResponseTimeText -Value ''
        Set-ControlText -Control $Ui.LastQueryBadgeText -Value 'Awaiting query'
        Update-FixNameButtonState -Ui $Ui -Device $null -ParentDevice $null
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
        return [Math]::Min(120, [Math]::Max(1, $minutes))
    }

    function Set-RoundingMinutes {
        param([hashtable]$Ui,[int]$Minutes)
        $Ui.RoundingTimeTextBox.Text = [Math]::Min(120, [Math]::Max(1, $Minutes)).ToString()
    }

    function Get-MaintenanceTypeOrDefault {
        param([string]$MaintenanceType,[string]$DeviceName)
        if (-not [string]::IsNullOrWhiteSpace($MaintenanceType)) { return $MaintenanceType.Trim() }
        if (-not [string]::IsNullOrWhiteSpace($DeviceName) -and $DeviceName.Trim() -match '(?i)^AO') { return 'Mobile Cart' }
        return 'General Rounding'
    }

    function Get-RoundingUrlForDevice {
        param([pscustomobject]$CurrentDevice,[hashtable]$RoundingByAssetTag)
        if (-not $CurrentDevice -or -not $CurrentDevice.AssetTag) { return $null }
        $key = $CurrentDevice.AssetTag.Trim().ToUpper()
        if ($RoundingByAssetTag.ContainsKey($key)) {
            return "https://devicerounding.nttdatanucleus.com/DeviceMaintenance/Index?DeviceId=$($RoundingByAssetTag[$key])"
        }
        return $null
    }

    function Load-RoundingAssetMap {
        param([string]$ResolvedXamlPath)
        $map = @{}
        $roundingMapPath = Get-RoundingMapPath -ResolvedXamlPath $ResolvedXamlPath
        if (-not (Test-Path -LiteralPath $roundingMapPath)) { return $map }
        try {
            foreach ($r in (Import-Csv -LiteralPath $roundingMapPath)) {
                $assetTag = [string]$r.'Asset Tag'
                $deviceId = [string]$r.SlNo
                if ([string]::IsNullOrWhiteSpace($assetTag) -or [string]::IsNullOrWhiteSpace($deviceId)) { continue }
                $map[$assetTag.Trim().ToUpper()] = $deviceId.Trim()
            }
        } catch {}
        return $map
    }

    function Update-ManualRoundButtonState {
        param([hashtable]$Ui,[pscustomobject]$CurrentDevice,[hashtable]$RoundingByAssetTag)
        $url = Get-RoundingUrlForDevice -CurrentDevice $CurrentDevice -RoundingByAssetTag $RoundingByAssetTag
        $Ui.ManualRoundButton.Tag = $url
        $Ui.ManualRoundButton.IsEnabled = -not [string]::IsNullOrWhiteSpace($url)
    }

    function Save-RoundingEvent {
        param([hashtable]$Ui,[pscustomobject]$CurrentDevice,[string]$ResolvedXamlPath)
        Ensure-OutputFolder -ResolvedXamlPath $ResolvedXamlPath
        $csvPath = Get-RoundingEventsPath -ResolvedXamlPath $ResolvedXamlPath
        $row = [pscustomobject]([ordered]@{
            Timestamp=(Get-Date).ToString('yyyy-MM-dd HH:mm:ss')
            AssetTag=$CurrentDevice.AssetTag; Name=$CurrentDevice.Name; Serial=$CurrentDevice.Serial
            City=$CurrentDevice.City; Location=$CurrentDevice.Location; Building=$CurrentDevice.Building
            Floor=$CurrentDevice.Floor; Room=$CurrentDevice.Room; CheckStatus=$Ui.CheckStatusComboBox.Text
            RoundingMinutes=(Get-RoundingMinutes -Ui $Ui); CableMgmtOK=$(if($Ui.ValidateCableCheckBox.IsChecked){'Yes'}else{'No'})
            CablingNeeded=$(if($Ui.CablingNeededCheckBox.IsChecked){'Yes'}else{'No'})
            LabelOK=$(if($Ui.LabelMonitorCheckBox.IsChecked){'Yes'}else{'No'}); CartOK=$(if($Ui.PhysicalCartCheckBox.IsChecked){'Yes'}else{'No'})
            PeripheralsOK=$(if($Ui.ValidatePeripheralsCheckBox.IsChecked){'Yes'}else{'No'})
            MaintenanceType=$Ui.MaintenanceTypeComboBox.Text; Department=$CurrentDevice.Department
            RoundingUrl=$Ui.ManualRoundButton.Tag; Comments=$Ui.CommentsTextBox.Text
            Rounded=$(if($script:ManualRoundUsed){'Yes'}else{'No'})
        })
        Add-RoundingCsvRow -Path $csvPath -Row $row
        $script:ManualRoundUsed = $false
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
            $row = [pscustomobject]([ordered]@{
                Timestamp=(Get-Date).ToString('yyyy-MM-dd HH:mm:ss')
                AssetTag=$item.AssetTag; Name=$item.HostName; Serial=''; City='Duncan'; Location=$item.Location; Building=$item.Building; Floor=$item.Floor; Room=$item.Room
                CheckStatus=$(if ([string]::IsNullOrWhiteSpace($item.Status) -or $item.Status -eq '-') { 'Complete' } else { $item.Status })
                RoundingMinutes=3; CableMgmtOK='No'; CablingNeeded='No'; LabelOK='No'; CartOK='No'; PeripheralsOK='No'
                MaintenanceType='Nearby'; Department=''; RoundingUrl=''; Comments='Saved from Nearby tab'; Rounded='No'
            })
            Add-RoundingCsvRow -Path $csvPath -Row $row
        }
        [System.Windows.MessageBox]::Show("Saved $($selectedRows.Count) nearby row(s) to:`n$csvPath", 'Nearby Save') | Out-Null
    }

    function Toggle-LocationEditMode {
        param([hashtable]$Ui,[bool]$IsEditing,[pscustomobject]$Inventory)
        $readOnlyControls = @($Ui.CityTextBox,$Ui.LocationTextBox,$Ui.BuildingTextBox,$Ui.FloorTextBox,$Ui.RoomTextBox,$Ui.DepartmentTextBox)
        $comboControls = @($Ui.CityComboBox,$Ui.LocationComboBox,$Ui.BuildingComboBox,$Ui.FloorComboBox,$Ui.RoomComboBox,$Ui.DepartmentComboBox)
        $visible = [System.Windows.Visibility]::Visible
        $collapsed = [System.Windows.Visibility]::Collapsed
        foreach ($control in $readOnlyControls) { if ($control) { $control.Visibility = if ($IsEditing) { $collapsed } else { $visible } } }
        foreach ($control in $comboControls) { if ($control) { $control.Visibility = if ($IsEditing) { $visible } else { $collapsed } } }
        $Ui.CancelEditLocationButton.Visibility = if ($IsEditing) { $visible } else { $collapsed }
        $Ui.EditLocationButton.Content = if ($IsEditing) { 'Save' } else { 'Edit Location' }
        if ($IsEditing -and $Inventory) {
            foreach ($pair in @(@($Ui.CityComboBox,$Ui.CityTextBox),@($Ui.LocationComboBox,$Ui.LocationTextBox),@($Ui.BuildingComboBox,$Ui.BuildingTextBox),@($Ui.FloorComboBox,$Ui.FloorTextBox),@($Ui.RoomComboBox,$Ui.RoomTextBox),@($Ui.DepartmentComboBox,$Ui.DepartmentTextBox))) { $pair[0].IsEditable = $true; $pair[0].Text = $pair[1].Text }
            Populate-LocationCombos -Ui $Ui -Inventory $Inventory
        }
    }

    function Save-LocationValues {
        param([hashtable]$Ui)

        Set-ControlText -Control $Ui.CityTextBox -Value $Ui.CityComboBox.Text
        Set-ControlText -Control $Ui.LocationTextBox -Value $Ui.LocationComboBox.Text
        Set-ControlText -Control $Ui.BuildingTextBox -Value $Ui.BuildingComboBox.Text
        Set-ControlText -Control $Ui.FloorTextBox -Value $Ui.FloorComboBox.Text
        Set-ControlText -Control $Ui.RoomTextBox -Value $Ui.RoomComboBox.Text
        Set-ControlText -Control $Ui.DepartmentTextBox -Value $Ui.DepartmentComboBox.Text
        if ($script:AppState) {
            $targets = @($script:AppState.SelectedSummaryParent,$script:AppState.SelectedSummaryDevice,$script:AppState.CurrentDevice) | Where-Object { $_ }
            foreach ($target in $targets) {
                $target.City = $Ui.CityTextBox.Text
                $target.Location = $Ui.LocationTextBox.Text
                $target.Building = $Ui.BuildingTextBox.Text
                $target.Floor = $Ui.FloorTextBox.Text
                $target.Room = $Ui.RoomTextBox.Text
                $target.Department = $Ui.DepartmentTextBox.Text
            }
            if ($script:AppState.Inventory) { Set-LocationValidationStyle -Ui $Ui -Inventory $script:AppState.Inventory }
        }
    }

    $resolvedXamlPath = (Resolve-Path -LiteralPath $XamlPath).Path
    $window = ConvertFrom-XamlFile -Path $resolvedXamlPath

    $ui = Get-NamedControls -Window $window -Names @(
        'SearchTextBox','QueryButton','PingButton','LiveDetailsButton','MonitorLabelButton',
        'MainTabControl','SystemTab','NearbyTab','SelectedDeviceText','DeviceOnlineText','DeviceOnlineDot','DeviceResponseTimeText','LastQueryBadgeText',
        'DetectedTypeDisplay','HostNameDisplay','AssetTagDisplay','SerialDisplay','ParentDisplay','RitmDisplay','RetireDisplay',
        'DetectedTypeTextBox','HostNameTextBox','AssetTagTextBox','SerialNumberTextBox','ParentTextBox','RitmTextBox','RetireDateTextBox','LastRoundedContainer','LastRoundedLabelText','LastRoundedText','LastRoundedAttentionBadge','LastRoundedAttentionText',
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

    $dataRoot = Join-Path (Split-Path -Parent $resolvedXamlPath) 'Data'
    $availableSites = Get-InventorySites -DataRoot $dataRoot
    $selectedSite = Select-InventorySite -Sites $availableSites
    if ($availableSites.Count -gt 0 -and $null -eq $selectedSite) { return }

    $siteFolderPath = if ($selectedSite) { $selectedSite.FolderPath } else { $null }
    $siteName = if ($selectedSite) { $selectedSite.Name } else { 'All Sites' }
    $inventory = Load-InventoryData -ResolvedXamlPath $resolvedXamlPath -SiteFolderPath $siteFolderPath
    $roundingByAssetTag = Load-RoundingAssetMap -ResolvedXamlPath $resolvedXamlPath
    $script:ManualRoundUsed = $false
    $roundingTimer = New-Object System.Windows.Threading.DispatcherTimer
    $roundingTimer.Interval = [TimeSpan]::FromSeconds(1)
    $script:RoundingStartTimeUtc = $null
    $script:RoundingBaseMinutes = 3
    $roundingTimer.Add_Tick({
        if (-not $script:RoundingStartTimeUtc) { return }
        $elapsed = [DateTime]::UtcNow - $script:RoundingStartTimeUtc
        $elapsedMinutes = [int][Math]::Floor($elapsed.TotalMinutes)
        $target = $script:RoundingBaseMinutes
        if ($elapsedMinutes -ge $script:RoundingBaseMinutes) { $target = [Math]::Min(120, $elapsedMinutes) }
        if ((Get-RoundingMinutes -Ui $ui) -lt $target) { Set-RoundingMinutes -Ui $ui -Minutes $target }
    })
    $script:AppState = [pscustomobject]@{ LastStatusMode='Found'; SampleData=$null; CurrentDevice=$null; CurrentQueryToken=''; Inventory=$inventory; SelectedSiteName=$siteName; SelectedSummaryDevice=$null; SelectedSummaryParent=$null }

    Clear-WindowData -Ui $ui
    Set-RoundingMinutes -Ui $ui -Minutes 3
    Increment-Fonts -Root $window
    Set-StatusMessage -Ui $ui -Mode 'Warning' -CustomText 'Ready. Enter a device and click Query.'
    Toggle-LocationEditMode -Ui $ui -IsEditing:$false

    $window.Title = "New Inventory Tool - $siteName"
    Set-ControlText -Control $ui.DataPathText -Value "Data: $dataRoot (Site: $siteName)"
    Set-ControlText -Control $ui.OutputPathText -Value "Output: $(Get-OutputFolder -ResolvedXamlPath $resolvedXamlPath)"

    foreach ($combo in @($ui.CityComboBox,$ui.LocationComboBox,$ui.BuildingComboBox,$ui.FloorComboBox,$ui.RoomComboBox,$ui.DepartmentComboBox)) {
        $combo.Items.Clear()
        Set-ControlText -Control $combo -Value ''
    }

    $ui.CityComboBox.Add_SelectionChanged({ if (-not $script:IsPopulatingLocationCombos) { Populate-LocationCombos -Ui $ui -Inventory $script:AppState.Inventory -ChangedLevel 'City' } })
    $ui.LocationComboBox.Add_SelectionChanged({ if (-not $script:IsPopulatingLocationCombos) { Populate-LocationCombos -Ui $ui -Inventory $script:AppState.Inventory -ChangedLevel 'Location' } })
    $ui.BuildingComboBox.Add_SelectionChanged({ if (-not $script:IsPopulatingLocationCombos) { Populate-LocationCombos -Ui $ui -Inventory $script:AppState.Inventory -ChangedLevel 'Building' } })
    $ui.FloorComboBox.Add_SelectionChanged({ if (-not $script:IsPopulatingLocationCombos) { Populate-LocationCombos -Ui $ui -Inventory $script:AppState.Inventory -ChangedLevel 'Floor' } })

    $ui.QueryButton.Add_Click({
        $searchTerm = $ui.SearchTextBox.Text
        $match = Find-InventoryMatch -SearchTerm $searchTerm -Inventory $script:AppState.Inventory
        if ($null -eq $match) {
            $script:AppState.CurrentDevice = $null
            $script:ManualRoundUsed = $false
            Update-ManualRoundButtonState -Ui $ui -CurrentDevice $null -RoundingByAssetTag $roundingByAssetTag
            Set-StatusMessage -Ui $ui -Mode 'Warning' -CustomText 'No matching device found'
            if (-not [string]::IsNullOrWhiteSpace($searchTerm)) {
                [System.Windows.MessageBox]::Show("No device was found for:`n$searchTerm", 'Device Not Found') | Out-Null
            }
            return
        }
        $script:AppState.CurrentDevice = $match
        $script:ManualRoundUsed = $false
        $maintenanceType = ''
        try { $maintenanceType = [string]$match.u_device_rounding } catch {}
        Set-ControlText -Control $ui.MaintenanceTypeComboBox -Value (Get-MaintenanceTypeOrDefault -MaintenanceType $maintenanceType -DeviceName $match.Name)
        Update-ManualRoundButtonState -Ui $ui -CurrentDevice $match -RoundingByAssetTag $roundingByAssetTag
        $script:RoundingBaseMinutes = 3
        Set-RoundingMinutes -Ui $ui -Minutes $script:RoundingBaseMinutes
        $script:RoundingStartTimeUtc = [DateTime]::UtcNow
        $roundingTimer.Start()
        $associated = Build-AssociatedDevices -Device $match -Inventory $script:AppState.Inventory
        $script:AppState.SampleData = [pscustomobject]@{ Device=$match; Associated=$associated; Nearby=@() }
        $script:AppState.CurrentQueryToken = [guid]::NewGuid().ToString('N')
        Set-PrimaryDeviceBindings -Ui $ui -Device $match -Inventory $inventory
        $ui.AssociatedDevicesDataGrid.ItemsSource = $associated
        $ui.NearbyDataGrid.ItemsSource = @()
        Set-ControlText -Control $ui.NearbyScopeSummaryText -Value 'Nearby disabled'
        Set-ControlText -Control $ui.DeviceOnlineText -Value 'Checking...'
        $ui.DeviceOnlineText.Foreground = New-Brush '#64748B'
        $ui.DeviceOnlineDot.Fill = New-Brush '#94A3B8'
        Set-ControlText -Control $ui.DeviceResponseTimeText -Value ''
        Set-ControlText -Control $ui.LastQueryBadgeText -Value "Queried $(Get-Date -Format 'HH:mm:ss')"
        Set-StatusMessage -Ui $ui -Mode 'Found'
        Start-OnlineStatusUpdateAsync -Ui $ui -HostName $match.Name
    })

    function Reset-RoundingFormForNextScan {
        param([hashtable]$Ui)
        foreach ($cb in @($Ui.ValidateCableCheckBox,$Ui.LabelMonitorCheckBox,$Ui.ValidatePeripheralsCheckBox,$Ui.CablingNeededCheckBox,$Ui.PhysicalCartCheckBox,$Ui.AddDeviceToTrackerCheckBox)) { $cb.IsChecked = $false }
        Set-ControlText -Control $Ui.CheckStatusComboBox -Value 'Complete'
        Set-ControlText -Control $Ui.MaintenanceTypeComboBox -Value 'General Rounding'
        Set-ControlText -Control $Ui.CommentsTextBox -Value ''
        Set-ControlText -Control $Ui.SearchTextBox -Value ''
        Set-RoundingMinutes -Ui $Ui -Minutes 3
        $Ui.SearchTextBox.Focus() | Out-Null
        $Ui.SearchTextBox.CaretIndex = $Ui.SearchTextBox.Text.Length
    }

    $ui.SearchTextBox.Add_TextChanged({
        if ([string]::IsNullOrWhiteSpace($ui.SearchTextBox.Text)) {
            $roundingTimer.Stop()
            $script:RoundingStartTimeUtc = $null
            $script:AppState.CurrentDevice = $null
            $script:AppState.SampleData = $null
            $script:AppState.CurrentQueryToken = ''
            $script:AppState.SelectedSummaryDevice = $null
            $script:AppState.SelectedSummaryParent = $null
            Clear-WindowData -Ui $ui
            Set-RoundingMinutes -Ui $ui -Minutes 3
            Update-ManualRoundButtonState -Ui $ui -CurrentDevice $null -RoundingByAssetTag $roundingByAssetTag
            Set-StatusMessage -Ui $ui -Mode 'Warning' -CustomText 'Ready. Enter a device and click Query.'
        }
    })

    $ui.SearchTextBox.Add_KeyDown({
        if ($_.Key -eq 'Return') {
            $ui.QueryButton.RaiseEvent((New-Object System.Windows.RoutedEventArgs([System.Windows.Controls.Button]::ClickEvent)))
        }
    })

    $ui.AddPeripheralButton.Add_Click({
        if (-not $script:AppState.CurrentDevice) { return }
        $parent = Resolve-ParentDevice -Device $script:AppState.CurrentDevice -Inventory $script:AppState.Inventory
        if (-not $parent) { $parent = $script:AppState.CurrentDevice }
        $changed = Show-AddPeripheralDialog -Ui $ui -ParentDevice $parent -Inventory $script:AppState.Inventory -ResolvedXamlPath $resolvedXamlPath
        if ($changed) {
            $associated = Build-AssociatedDevices -Device $script:AppState.CurrentDevice -Inventory $script:AppState.Inventory
            $ui.AssociatedDevicesDataGrid.ItemsSource = $associated
            $script:AppState.SampleData = [pscustomobject]@{ Device=$script:AppState.CurrentDevice; Associated=$associated; Nearby=@() }
        }
    })
    $ui.RemovePeripheralButton.Add_Click({
        if (-not $script:AppState.CurrentDevice) { return }
        $changed = Remove-SelectedPeripheralAssociations -Ui $ui -Inventory $script:AppState.Inventory -ResolvedXamlPath $resolvedXamlPath
        if ($changed) {
            $associated = Build-AssociatedDevices -Device $script:AppState.CurrentDevice -Inventory $script:AppState.Inventory
            $ui.AssociatedDevicesDataGrid.ItemsSource = $associated
            $script:AppState.SampleData = [pscustomobject]@{ Device=$script:AppState.CurrentDevice; Associated=$associated; Nearby=@() }
            Set-SelectedSummaryDevice -Ui $ui -Device $script:AppState.CurrentDevice -Inventory $script:AppState.Inventory
        }
    })
    $ui.ValidateAssociatedButton.Add_Click({
        if (-not $script:AppState.CurrentDevice) { return }
        $parent = Resolve-ParentDevice -Device $script:AppState.CurrentDevice -Inventory $script:AppState.Inventory
        if (-not $parent) { $parent = $script:AppState.CurrentDevice }
        Validate-AssociatedDevices -Ui $ui -ParentDevice $parent -Inventory $script:AppState.Inventory -ResolvedXamlPath $resolvedXamlPath
    })
    $ui.PingButton.Add_Click({
        $target = $null
        if ($script:AppState.CurrentDevice -and -not [string]::IsNullOrWhiteSpace($script:AppState.CurrentDevice.Name)) { $target = $script:AppState.CurrentDevice.Name }
        elseif (-not [string]::IsNullOrWhiteSpace($ui.SearchTextBox.Text)) { $target = $ui.SearchTextBox.Text.Trim() }
        try {
            $pingResult = Invoke-DevicePing -ComputerName $target -DataRoot $dataRoot
            Start-ContinuousPingWindow -Target $(if ($pingResult.IpAddress -and $pingResult.IpAddress -ne 'Unknown') { $pingResult.IpAddress } else { $target })
            Show-PingResultDialog -PingResult $pingResult
            Set-StatusMessage -Ui $ui -Mode 'PingComplete' -CustomText $(if ($pingResult.Success) { "Ping complete: $($pingResult.ResponseTime)" } else { 'Ping failed' })
        } catch {
            [System.Windows.MessageBox]::Show($_.Exception.Message, 'Ping') | Out-Null
        }
    })
    $ui.LiveDetailsButton.Add_Click({ [System.Windows.MessageBox]::Show('Live Details button clicked.', 'Live Details') | Out-Null })
    $ui.MonitorLabelButton.Add_Click({
        $device = $script:AppState.CurrentDevice
        if (-not $device -and -not [string]::IsNullOrWhiteSpace($ui.SearchTextBox.Text)) {
            $device = Find-InventoryMatch -SearchTerm $ui.SearchTextBox.Text -Inventory $script:AppState.Inventory
        }
        $parent = if ($device) { Resolve-ParentDevice -Device $device -Inventory $script:AppState.Inventory } else { $null }
        if (-not $parent) { $parent = $device }
        if (-not $parent) {
            [System.Windows.MessageBox]::Show('No parent device available to show.', 'Monitor Label') | Out-Null
            return
        }
        Show-MonitorLabelDialog -Ui $ui -ParentDevice $parent
    })
    $ui.FixNameButton.Add_Click({
        $device = $script:AppState.SelectedSummaryDevice
        $parent = $script:AppState.SelectedSummaryParent
        if (-not $device) { return }
        if (-not $parent) { [System.Windows.MessageBox]::Show('No parent computer found for this device. Scan a device with a valid parent first.', 'Fix Name') | Out-Null; return }
        $expected = Get-ExpectedDeviceName -Device $device -ParentDevice $parent
        if ([string]::IsNullOrWhiteSpace($expected)) { [System.Windows.MessageBox]::Show('Could not compute the expected name for this device type.', 'Fix Name') | Out-Null; return }
        if ($device.Name -and $device.Name.Trim().ToUpper() -eq $expected.Trim().ToUpper()) {
            Update-FixNameButtonState -Ui $ui -Device $device -ParentDevice $parent
            return
        }
        $oldName = $device.Name
        $device.Name = $expected
        Update-InventoryDeviceName -Inventory $script:AppState.Inventory -Device $device -NewName $expected
        Add-CmdbAssociationUpdate -ResolvedXamlPath $resolvedXamlPath -Candidate $device -OldParent $device.Parent -NewParent $device.Parent -OldName $oldName -NewName $expected -Action 'Link' -Inventory $script:AppState.Inventory
        Build-InventoryIndices -Inventory $script:AppState.Inventory
        Set-SelectedSummaryDevice -Ui $ui -Device $device -Inventory $script:AppState.Inventory
        if ($script:AppState.CurrentDevice) {
            $associated = Build-AssociatedDevices -Device $script:AppState.CurrentDevice -Inventory $script:AppState.Inventory
            $ui.AssociatedDevicesDataGrid.ItemsSource = $associated
            $script:AppState.SampleData = [pscustomobject]@{ Device=$script:AppState.CurrentDevice; Associated=$associated; Nearby=@() }
        }
    })
    $ui.EditLocationButton.Add_Click([System.Windows.RoutedEventHandler]{
        param($sender,$e)
        try {
            if ([string]$ui.EditLocationButton.Content -eq 'Save') {
                Save-LocationValues -Ui $ui
                Toggle-LocationEditMode -Ui $ui -IsEditing:$false
            }
            else {
                Toggle-LocationEditMode -Ui $ui -IsEditing:$true -Inventory $script:AppState.Inventory
            }
        } catch {
            [System.Windows.MessageBox]::Show($_.Exception.Message, 'Edit Location') | Out-Null
        }
    })
    $ui.CancelEditLocationButton.Add_Click([System.Windows.RoutedEventHandler]{
        param($sender,$e)
        try {
            Toggle-LocationEditMode -Ui $ui -IsEditing:$false
            Set-LocationValidationStyle -Ui $ui -Inventory $script:AppState.Inventory
        } catch {
            [System.Windows.MessageBox]::Show($_.Exception.Message, 'Edit Location') | Out-Null
        }
    })
    $ui.RoundingTimeUpButton.Add_Click({ Set-RoundingMinutes -Ui $ui -Minutes ((Get-RoundingMinutes -Ui $ui) + 1); $script:RoundingBaseMinutes = (Get-RoundingMinutes -Ui $ui); $roundingTimer.Stop(); $script:RoundingStartTimeUtc = $null })
    $ui.RoundingTimeDownButton.Add_Click({ Set-RoundingMinutes -Ui $ui -Minutes ((Get-RoundingMinutes -Ui $ui) - 1); $script:RoundingBaseMinutes = (Get-RoundingMinutes -Ui $ui); $roundingTimer.Stop(); $script:RoundingStartTimeUtc = $null })
    $ui.CheckCompleteButton.Add_Click({
        $checkBoxes = @($ui.ValidateCableCheckBox,$ui.LabelMonitorCheckBox,$ui.ValidatePeripheralsCheckBox,$ui.PhysicalCartCheckBox)
        $enabled = @($checkBoxes | Where-Object { $_.IsEnabled })
        $allChecked = $enabled.Count -gt 0 -and (@($enabled | Where-Object { -not $_.IsChecked }).Count -eq 0)
        if ($allChecked) {
            foreach ($cb in $checkBoxes) { $cb.IsChecked = $false }
        } else {
            foreach ($cb in $enabled) { $cb.IsChecked = $true }
        }
        Set-ControlText -Control $ui.CheckStatusComboBox -Value 'Complete'
    })
    $ui.SaveEventButton.Add_Click({
        if (-not $script:AppState.CurrentDevice) { return }
        $roundingTimer.Stop()
        $script:RoundingStartTimeUtc = $null
        if ($ui.MaintenanceTypeComboBox.Text -eq 'Excluded') {
            [System.Windows.MessageBox]::Show("This device is marked as Excluded. Enable 'Excluded' to log rounding.","Save Event") | Out-Null
            return
        }
        Save-RoundingEvent -Ui $ui -CurrentDevice $script:AppState.CurrentDevice -ResolvedXamlPath $resolvedXamlPath
        $script:RoundingBaseMinutes = 3
        Reset-RoundingFormForNextScan -Ui $ui
    })
    $ui.ManualRoundButton.Add_Click({
        if ([string]::IsNullOrWhiteSpace($ui.ManualRoundButton.Tag)) {
            [System.Windows.MessageBox]::Show("No rounding URL found for this device.","Manual Round") | Out-Null
            return
        }
        $script:ManualRoundUsed = $true
        Start-Process -FilePath $ui.ManualRoundButton.Tag
    })
    $ui.RebuildNearbyButton.Add_Click({ Set-ControlText -Control $ui.NearbyScopeSummaryText -Value 'Nearby disabled' })
    $ui.ClearNearbyButton.Add_Click({ $ui.NearbyDataGrid.ItemsSource = @(); Set-ControlText -Control $ui.NearbyScopeSummaryText -Value 'Nearby disabled' })
    $ui.PingAllButton.Add_Click({ Set-ControlText -Control $ui.NearbyScopeSummaryText -Value 'Nearby disabled' })
    $ui.NearbySaveButton.Add_Click({ [System.Windows.MessageBox]::Show('Nearby logic is currently disabled.', 'Nearby Disabled') | Out-Null })
    $ui.AssociatedDevicesDataGrid.Add_MouseDoubleClick({
        $row = $ui.AssociatedDevicesDataGrid.SelectedItem
        if (-not $row) { return }
        $device = $null
        if ($row.PSObject.Properties.Name -contains 'Device') { $device = $row.Device }
        if (-not $device) {
            $device = [pscustomobject]@{
                DetectedType=$row.Type; Name=$row.Name; AssetTag=$row.AssetTag; Serial=$row.Serial; Parent='(n/a)'; RITM=$row.RITM; RetireDate=$row.Retire
                LastRounded=''; City=''; Location=''; Building=''; Floor=''; Room=''; Department=''
            }
        }
        Set-SelectedSummaryDevice -Ui $ui -Device $device -Inventory $script:AppState.Inventory
    })

    $ui.AssociatedDevicesDataGrid.AddHandler([System.Windows.Documents.Hyperlink]::RequestNavigateEvent, [System.Windows.Navigation.RequestNavigateEventHandler]{
        param($sender,$e)
        if ($e.Uri) { Start-Process $e.Uri.AbsoluteUri }
        $e.Handled = $true
    })

    $ui.MainTabControl.Add_SelectionChanged({
        if ($window.IsLoaded) {
            if ($script:AppState.LastStatusMode -eq 'PingComplete') { Set-StatusMessage -Ui $ui -Mode 'PingComplete' }
            else { Set-StatusMessage -Ui $ui -Mode 'Found' }
        }
    })

    $ui.ValidatePeripheralsCheckBox.IsChecked = $false
    $ui.LabelMonitorCheckBox.IsChecked = $false
    $ui.ValidateCableCheckBox.IsChecked = $false
    $ui.CablingNeededCheckBox.IsChecked = $false
    $ui.PhysicalCartCheckBox.IsChecked = $false
    $ui.AddDeviceToTrackerCheckBox.IsChecked = $false
    $ui.ManualRoundButton.IsEnabled = $false

    [void]$window.ShowDialog()
}
catch {
    Show-StartupError -Exception $_.Exception -ScriptPath $PSCommandPath
    throw
}
