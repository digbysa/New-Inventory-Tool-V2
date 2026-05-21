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
            $Ui.LastRoundedText.Text = ''
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
        foreach ($name in $Names) {
            if ($Row.PSObject.Properties.Name -contains $name) {
                $value = $Row.$name
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
        return [pscustomobject]@{ DataRoot=$dataRoot; SiteFolderPath=$SiteFolderPath; Computers=$computers; Monitors=$monitors; Carts=$carts; Mics=$mics; Scanners=$scanners; Locations=$locations }
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

    function Find-InventoryMatch {
        param([string]$SearchTerm,[pscustomobject]$Inventory)
        $term = $SearchTerm.Trim().ToUpper()
        if ([string]::IsNullOrWhiteSpace($term)) { return $null }

        $isComputerName = $term -match '^(PC|LD|TD|AO)'
        $isAssetTag = $term -match '^(HSS|C)'

        function Find-InCollection {
            param([object[]]$Rows,[string[]]$Fields,[string]$DetectedType)
            foreach ($row in $Rows) {
                foreach ($field in $Fields) {
                    $candidate = Get-FieldValue -Row $row -Names @($field)
                    if (-not [string]::IsNullOrWhiteSpace($candidate) -and $candidate.Trim().ToUpper() -eq $term) {
                        return (ConvertTo-DeviceRecord -Row $row -DetectedType $DetectedType)
                    }
                }
            }
            return $null
        }

        if ($isComputerName) {
            $match = Find-InCollection -Rows $Inventory.Computers -Fields @('name') -DetectedType 'Computer'
            if ($match) { return $match }
        }

        if ($isAssetTag) {
            $match = Find-InCollection -Rows $Inventory.Computers -Fields @('asset_tag') -DetectedType 'Computer'
            if ($match) { return $match }
            $match = Find-InCollection -Rows $Inventory.Monitors -Fields @('asset_tag') -DetectedType 'Monitor'
            if ($match) { return $match }
            $match = Find-InCollection -Rows $Inventory.Carts -Fields @('asset_tag') -DetectedType 'Cart'
            if ($match) { return $match }
            $match = Find-InCollection -Rows $Inventory.Mics -Fields @('asset_tag') -DetectedType 'Mic'
            if ($match) { return $match }
            $match = Find-InCollection -Rows $Inventory.Scanners -Fields @('asset_tag') -DetectedType 'Scanner'
            if ($match) { return $match }
        }

        foreach ($set in @(
            @{Rows=$Inventory.Computers; Fields=@('name','asset_tag','serial_number'); Type='Computer'},
            @{Rows=$Inventory.Monitors; Fields=@('name','asset_tag','serial_number'); Type='Monitor'},
            @{Rows=$Inventory.Carts; Fields=@('name','asset_tag','serial_number'); Type='Cart'},
            @{Rows=$Inventory.Mics; Fields=@('name','asset_tag','serial_number'); Type='Mic'},
            @{Rows=$Inventory.Scanners; Fields=@('name','asset_tag','serial_number'); Type='Scanner'}
        )) {
            $match = Find-InCollection -Rows $set.Rows -Fields $set.Fields -DetectedType $set.Type
            if ($match) { return $match }
        }
        return $null
    }


    function Resolve-ParentDevice {
        param([pscustomobject]$Device,[pscustomobject]$Inventory)
        if (-not $Device) { return $null }
        if ($Device.DetectedType -eq 'Computer' -and $Device.Name -match '^(LD|PC|TD|AO)') { return $Device }

        $tokens = @(@($Device.Parent) | Where-Object { -not [string]::IsNullOrWhiteSpace($_) } | ForEach-Object { $_.Trim().ToUpper() })
        if (-not $tokens -or $tokens.Count -eq 0) { return $null }

        foreach ($row in $Inventory.Computers) {
            $record = ConvertTo-DeviceRecord -Row $row -DetectedType 'Computer'
            $name = if ($record.Name) { $record.Name.Trim().ToUpper() } else { '' }
            if ($name -notmatch '^(LD|PC|TD|AO)') { continue }
            $candidates = @($record.Name,$record.AssetTag,$record.Serial) | Where-Object { -not [string]::IsNullOrWhiteSpace($_) } | ForEach-Object { $_.Trim().ToUpper() }
            foreach ($token in $tokens) {
                if ($candidates -contains $token) { return $record }
            }
        }
        return $null
    }
    function Build-AssociatedDevices {
        param([pscustomobject]$Device,[pscustomobject]$Inventory)
        $parentDevice = Resolve-ParentDevice -Device $Device -Inventory $Inventory
        $effectiveParent = if ($parentDevice) { $parentDevice } else { $Device }
        $queryRole = if ($parentDevice) { 'Child' } else { 'Parent' }

        $results = @([pscustomobject]@{ Role='Parent'; Type=$effectiveParent.DetectedType; Name=$effectiveParent.Name; AssetTag=$effectiveParent.AssetTag; Serial=$effectiveParent.Serial; RITM=$effectiveParent.RITM; Retire=(Format-DateLong $effectiveParent.RetireDate); CmdbUrl=(Get-CmdbLink -DeviceType $effectiveParent.DetectedType -AssetTag $effectiveParent.AssetTag) })
        $childrenByParent = @{}
        foreach ($collectionName in @('Monitors','Carts','Mics','Scanners')) {
            $collection = $Inventory.$collectionName
            if (-not $collection) { continue }
            foreach ($row in $collection) {
                $token = (Get-FieldValue -Row $row -Names @('u_parent_asset','Parent')).Trim()
                if ([string]::IsNullOrWhiteSpace($token)) { continue }
                $key = $token.ToUpper()
                if (-not $childrenByParent.ContainsKey($key)) { $childrenByParent[$key] = @() }
                $childrenByParent[$key] += ,$row
            }
        }
        function Add-AssociatedRecord {
            param($Role,$Type,$Row)
            $record = [pscustomobject]@{ Role=$Role; Type=$Type; Name=(Get-FieldValue -Row $Row -Names @('name')); AssetTag=(Get-FieldValue -Row $Row -Names @('asset_tag')); Serial=(Get-FieldValue -Row $Row -Names @('serial_number')); RITM=(Extract-Ritm (Get-FieldValue -Row $Row -Names @('po_number'))); Retire=(Format-DateLong (Get-FieldValue -Row $Row -Names @('u_scheduled_retirement'))); CmdbUrl=(Get-CmdbLink -DeviceType $Type -AssetTag (Get-FieldValue -Row $Row -Names @('asset_tag'))) }
            $results = @($results + $record)
        }
        $parentTokens = @()
        $parentTokens += @($effectiveParent.AssetTag,$effectiveParent.Name,$effectiveParent.Serial)
        if (-not [string]::IsNullOrWhiteSpace($Device.Parent)) { $parentTokens += $Device.Parent }
        $parentTokens = @($parentTokens | Where-Object { -not [string]::IsNullOrWhiteSpace($_) } | ForEach-Object { $_.Trim().ToUpper() } | Select-Object -Unique)

        $addedChildAssetTags = @{}
        foreach ($token in $parentTokens) {
            if (-not $childrenByParent.ContainsKey($token)) { continue }
            foreach ($row in $childrenByParent[$token]) {
                $childAssetTag = (Get-FieldValue -Row $row -Names @('asset_tag')).Trim().ToUpper()
                if (-not [string]::IsNullOrWhiteSpace($childAssetTag) -and $addedChildAssetTags.ContainsKey($childAssetTag)) { continue }
                $type = if ($Inventory.Carts -contains $row) { 'Cart' } elseif ($Inventory.Mics -contains $row) { 'Mic' } elseif ($Inventory.Scanners -contains $row) { 'Scanner' } else { 'Monitor' }
                $role = if ($Device.AssetTag -and ($childAssetTag -eq $Device.AssetTag.Trim().ToUpper())) { $queryRole } else { 'Child' }
                Add-AssociatedRecord -Role $role -Type $type -Row $row
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

    function Show-AddPeripheralDialog {
        param([hashtable]$Ui,[pscustomobject]$ParentDevice,[pscustomobject]$Inventory)
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
        $txt = New-Object System.Windows.Controls.TextBox
        $txt.MinWidth = 320
        $txt.Margin = '0,0,0,6'
        $panel.Children.Add($txt) | Out-Null
        $hint = New-Object System.Windows.Controls.TextBlock -Property @{ Text='Press Enter to search by cart name, asset tag, or serial number.'; Foreground='#64748B'; Margin='0,0,0,8' }
        $panel.Children.Add($hint) | Out-Null
        $preview = New-Object System.Windows.Controls.TextBlock -Property @{ Text=''; Margin='0,0,0,8'; TextWrapping='Wrap' }
        $panel.Children.Add($preview) | Out-Null
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
            if (-not $result -or -not $result.Candidate) {
                $preview.Text = if ($result -and $result.NormalizedInput) { 'Type: (not found)' } else { '' }
                $addButton.IsEnabled = $false
                return
            }
            $c = $result.Candidate
            $preview.Text = "Type: $($c.DetectedType)`nName: $($c.Name)`nAsset Tag: $($c.AssetTag)`nSerial: $($c.Serial)`nRITM: $($c.RITM)`nRetire: $($c.RetireDate)`nParent: $($c.Parent)"
            $addButton.IsEnabled = $true
        }

        $txt.Add_TextChanged({ $lookupResult = $null; $preview.Text = ''; $addButton.IsEnabled = $false })
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
            foreach ($collectionName in @('Monitors','Carts','Mics','Scanners')) {
                foreach ($row in $Inventory.$collectionName) {
                    $at = Get-FieldValue -Row $row -Names @('asset_tag')
                    if ($at -and $target.AssetTag -and $at.Trim().ToUpper() -eq $target.AssetTag.Trim().ToUpper()) {
                        $row.u_parent_asset = $ParentDevice.AssetTag
                    }
                }
            }
            $window.DialogResult = $true
            $window.Close()
        })
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


    function Set-PrimaryDeviceBindings {
        param([hashtable]$Ui,[pscustomobject]$Device,[pscustomobject]$Inventory)
        $parentDevice = Resolve-ParentDevice -Device $Device -Inventory $Inventory
        $summaryDevice = if ($parentDevice) { $parentDevice } else { $Device }

        $Ui.SelectedDeviceText.Text = $Device.Name
        Set-DisplayText -Ui $Ui -BaseName 'DetectedType' -Value $Device.DetectedType
        Set-DisplayText -Ui $Ui -BaseName 'HostName' -Value $Device.Name
        Set-DisplayText -Ui $Ui -BaseName 'AssetTag' -Value $Device.AssetTag
        Set-DisplayText -Ui $Ui -BaseName 'Serial' -Value $Device.Serial
        Set-DisplayText -Ui $Ui -BaseName 'Parent' -Value $Device.Parent
        Set-DisplayText -Ui $Ui -BaseName 'Ritm' -Value $Device.RITM
        Set-DisplayText -Ui $Ui -BaseName 'Retire' -Value (Format-DateLong $Device.RetireDate)
        Set-LastRoundedDisplay -Ui $Ui -LastRoundedRaw $summaryDevice.LastRounded
        $Ui.CityTextBox.Text = $summaryDevice.City
        $Ui.LocationTextBox.Text = $summaryDevice.Location
        $Ui.BuildingTextBox.Text = $summaryDevice.Building
        $Ui.FloorTextBox.Text = $summaryDevice.Floor
        $Ui.RoomTextBox.Text = $summaryDevice.Room
        $Ui.DepartmentTextBox.Text = $summaryDevice.Department
    }

    function Start-QueryDataPopulationAsync {
        param([hashtable]$Ui,[pscustomobject]$Device,[pscustomobject]$Inventory,[string]$QueryToken)
        [System.Threading.Tasks.Task]::Run([Action]{
            $associated = Build-AssociatedDevices -Device $Device -Inventory $Inventory
            $Ui.MainTabControl.Dispatcher.BeginInvoke([Action]{
                if ($script:AppState.CurrentQueryToken -ne $QueryToken) { return }
                $Ui.AssociatedDevicesDataGrid.ItemsSource = $associated
                $Ui.NearbyDataGrid.ItemsSource = @()
                $Ui.NearbyScopeSummaryText.Text = 'Nearby disabled'
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
        $Ui.LastQueryBadgeText.Text = "Queried $(Get-Date -Format 'HH:mm')"
        $Ui.NearbyScopeSummaryText.Text = 'Nearby disabled'
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
        $Ui.SelectedDeviceText.Text = ''
        $Ui.LastRoundedLabelText.Foreground = New-Brush '#64748B'
        $Ui.LastRoundedText.Text = ''
        $Ui.LastRoundedText.Foreground = New-Brush '#64748B'
        $Ui.LastRoundedContainer.Background = New-Brush '#F8FAFC'
        $Ui.LastRoundedContainer.BorderBrush = New-Brush '#D9E1EA'
        $Ui.LastRoundedAttentionBadge.Visibility = 'Collapsed'
        foreach ($box in @($Ui.CityTextBox,$Ui.LocationTextBox,$Ui.BuildingTextBox,$Ui.FloorTextBox,$Ui.RoomTextBox,$Ui.DepartmentTextBox)) { $box.Text = '' }
        $Ui.NearbyScopeSummaryText.Text = 'Nearby disabled'
        $Ui.AssociatedDevicesDataGrid.ItemsSource = @()
        $Ui.NearbyDataGrid.ItemsSource = @()
        $Ui.DeviceOnlineText.Text = 'Ready'
        $Ui.DeviceOnlineText.Foreground = New-Brush '#64748B'
        $Ui.DeviceOnlineDot.Fill = New-Brush '#94A3B8'
        $Ui.DeviceResponseTimeText.Text = ''
        $Ui.LastQueryBadgeText.Text = 'Awaiting query'
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
    $script:AppState = [pscustomobject]@{ LastStatusMode='Found'; SampleData=$null; CurrentDevice=$null; CurrentQueryToken=''; Inventory=$inventory; SelectedSiteName=$siteName }

    Clear-WindowData -Ui $ui
    Increment-Fonts -Root $window
    Set-StatusMessage -Ui $ui -Mode 'Warning' -CustomText 'Ready. Enter a device and click Query.'
    Toggle-LocationEditMode -Ui $ui -IsEditing:$false

    $window.Title = "New Inventory Tool - $siteName"
    $ui.DataPathText.Text = "Data: $dataRoot (Site: $siteName)"
    $ui.OutputPathText.Text = "Output: $(Get-OutputFolder -ResolvedXamlPath $resolvedXamlPath)"

    foreach ($combo in @($ui.CityComboBox,$ui.LocationComboBox,$ui.BuildingComboBox,$ui.FloorComboBox,$ui.RoomComboBox,$ui.DepartmentComboBox)) {
        $combo.Items.Clear()
        $combo.Text = ''
    }

    $ui.QueryButton.Add_Click({
        $searchTerm = $ui.SearchTextBox.Text
        $match = Find-InventoryMatch -SearchTerm $searchTerm -Inventory $script:AppState.Inventory
        if ($null -eq $match) {
            Set-StatusMessage -Ui $ui -Mode 'Warning' -CustomText 'No matching device found'
            if (-not [string]::IsNullOrWhiteSpace($searchTerm)) {
                [System.Windows.MessageBox]::Show("No device was found for:`n$searchTerm", 'Device Not Found') | Out-Null
            }
            return
        }
        $script:AppState.CurrentDevice = $match
        $associated = Build-AssociatedDevices -Device $match -Inventory $script:AppState.Inventory
        $script:AppState.SampleData = [pscustomobject]@{ Device=$match; Associated=$associated; Nearby=@() }
        $script:AppState.CurrentQueryToken = [guid]::NewGuid().ToString('N')
        Set-PrimaryDeviceBindings -Ui $ui -Device $match -Inventory $inventory
        $ui.AssociatedDevicesDataGrid.ItemsSource = $associated
        $ui.NearbyDataGrid.ItemsSource = @()
        $ui.NearbyScopeSummaryText.Text = 'Nearby disabled'
        $ui.DeviceOnlineText.Text = 'Checking...'
        $ui.DeviceOnlineText.Foreground = New-Brush '#64748B'
        $ui.DeviceOnlineDot.Fill = New-Brush '#94A3B8'
        $ui.DeviceResponseTimeText.Text = ''
        $ui.LastQueryBadgeText.Text = "Queried $(Get-Date -Format 'HH:mm:ss')"
        Set-StatusMessage -Ui $ui -Mode 'Found'
        Start-OnlineStatusUpdateAsync -Ui $ui -HostName $match.Name
    })

    $ui.SearchTextBox.Add_TextChanged({
        if ([string]::IsNullOrWhiteSpace($ui.SearchTextBox.Text)) {
            $script:AppState.CurrentDevice = $null
            $script:AppState.SampleData = $null
            $script:AppState.CurrentQueryToken = ''
            Clear-WindowData -Ui $ui
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
        $changed = Show-AddPeripheralDialog -Ui $ui -ParentDevice $parent -Inventory $script:AppState.Inventory
        if ($changed) {
            $associated = Build-AssociatedDevices -Device $script:AppState.CurrentDevice -Inventory $script:AppState.Inventory
            $ui.AssociatedDevicesDataGrid.ItemsSource = $associated
            $script:AppState.SampleData = [pscustomobject]@{ Device=$script:AppState.CurrentDevice; Associated=$associated; Nearby=@() }
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
    $ui.RebuildNearbyButton.Add_Click({ $ui.NearbyScopeSummaryText.Text = 'Nearby disabled' })
    $ui.ClearNearbyButton.Add_Click({ $ui.NearbyDataGrid.ItemsSource = @(); $ui.NearbyScopeSummaryText.Text = 'Nearby disabled' })
    $ui.PingAllButton.Add_Click({ $ui.NearbyScopeSummaryText.Text = 'Nearby disabled' })
    $ui.NearbySaveButton.Add_Click({ [System.Windows.MessageBox]::Show('Nearby logic is currently disabled.', 'Nearby Disabled') | Out-Null })
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
