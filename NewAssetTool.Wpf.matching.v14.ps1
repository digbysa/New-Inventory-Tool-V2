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
        $rawXaml = Get-Content -LiteralPath $Path -Raw -Encoding UTF8
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

    function Find-VisualChildByType {
        param(
            [Parameter(Mandatory=$true)][System.Object]$Root,
            [Parameter(Mandatory=$true)][Type]$Type
        )

        if ($null -eq $Root) { return $null }
        if ($Root -is $Type) { return $Root }
        if (-not ($Root -is [System.Windows.Media.Visual] -or $Root -is [System.Windows.Media.Media3D.Visual3D])) { return $null }

        $count = [System.Windows.Media.VisualTreeHelper]::GetChildrenCount($Root)
        for ($i = 0; $i -lt $count; $i++) {
            $found = Find-VisualChildByType -Root ([System.Windows.Media.VisualTreeHelper]::GetChild($Root, $i)) -Type $Type
            if ($null -ne $found) { return $found }
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



    function Set-WindowIconFromFile {
        param(
            [Parameter(Mandatory=$true)][System.Windows.Window]$Window,
            [Parameter(Mandatory=$true)][string]$ResolvedXamlPath
        )

        $iconPath = Join-Path (Split-Path -Parent $ResolvedXamlPath) 'icon.ico'
        if (-not (Test-Path -LiteralPath $iconPath)) { return }

        $iconUri = New-Object System.Uri($iconPath, [System.UriKind]::Absolute)
        $Window.Icon = [System.Windows.Media.Imaging.BitmapFrame]::Create($iconUri)
    }

    function New-CalendarWindowIcon {
        $group = New-Object System.Windows.Media.DrawingGroup
        $outline = New-Object System.Windows.Media.GeometryDrawing
        $outline.Brush = New-Brush '#FFFFFF'
        $outline.Pen = New-Object System.Windows.Media.Pen((New-Brush '#0F5EA8'), 1.4)
        $outline.Geometry = [System.Windows.Media.Geometry]::Parse('M 4,5 L 12,5 L 12,13 L 4,13 Z')
        $group.Children.Add($outline) | Out-Null

        $header = New-Object System.Windows.Media.GeometryDrawing
        $header.Brush = New-Brush '#0EA5E9'
        $header.Geometry = [System.Windows.Media.Geometry]::Parse('M 4,5 L 12,5 L 12,7.2 L 4,7.2 Z')
        $group.Children.Add($header) | Out-Null

        $rings = New-Object System.Windows.Media.GeometryDrawing
        $rings.Brush = New-Brush '#1F2A44'
        $rings.Geometry = [System.Windows.Media.Geometry]::Parse('M 5.5,3.5 L 5.5,6 M 10.5,3.5 L 10.5,6')
        $rings.Pen = New-Object System.Windows.Media.Pen((New-Brush '#1F2A44'), 1.2)
        $group.Children.Add($rings) | Out-Null

        $dots = New-Object System.Windows.Media.GeometryDrawing
        $dots.Brush = New-Brush '#16A34A'
        $dots.Geometry = [System.Windows.Media.Geometry]::Parse('M 5.5,8.4 L 6.8,8.4 L 6.8,9.7 L 5.5,9.7 Z M 8.8,8.4 L 10.1,8.4 L 10.1,9.7 L 8.8,9.7 Z M 5.5,10.7 L 6.8,10.7 L 6.8,12 L 5.5,12 Z M 8.8,10.7 L 10.1,10.7 L 10.1,12 L 8.8,12 Z')
        $group.Children.Add($dots) | Out-Null

        $image = New-Object System.Windows.Media.DrawingImage($group)
        $image.Freeze()
        return $image
    }

    function Set-RoundedButtonTemplate {
        param([System.Windows.Controls.Button]$Button,[double]$CornerRadius = 8)
        $template = New-Object System.Windows.Controls.ControlTemplate([System.Windows.Controls.Button])
        $border = New-Object System.Windows.FrameworkElementFactory([System.Windows.Controls.Border])
        $border.SetValue([System.Windows.Controls.Border]::BackgroundProperty, (New-Object System.Windows.TemplateBindingExtension([System.Windows.Controls.Button]::BackgroundProperty)))
        $border.SetValue([System.Windows.Controls.Border]::BorderBrushProperty, (New-Object System.Windows.TemplateBindingExtension([System.Windows.Controls.Button]::BorderBrushProperty)))
        $border.SetValue([System.Windows.Controls.Border]::BorderThicknessProperty, (New-Object System.Windows.TemplateBindingExtension([System.Windows.Controls.Button]::BorderThicknessProperty)))
        $border.SetValue([System.Windows.Controls.Border]::CornerRadiusProperty, (New-Object System.Windows.CornerRadius($CornerRadius)))
        $border.SetValue([System.Windows.Controls.Border]::PaddingProperty, (New-Object System.Windows.TemplateBindingExtension([System.Windows.Controls.Button]::PaddingProperty)))
        $presenter = New-Object System.Windows.FrameworkElementFactory([System.Windows.Controls.ContentPresenter])
        $presenter.SetValue([System.Windows.Controls.ContentPresenter]::HorizontalAlignmentProperty, [System.Windows.HorizontalAlignment]::Center)
        $presenter.SetValue([System.Windows.Controls.ContentPresenter]::VerticalAlignmentProperty, [System.Windows.VerticalAlignment]::Center)
        $border.AppendChild($presenter)
        $template.VisualTree = $border
        $Button.Template = $template
    }

    function Get-OutputFolder { param([string]$ResolvedXamlPath) Join-Path (Split-Path -Parent $ResolvedXamlPath) 'Output' }
    function Get-RoundingEventsPath { param([string]$ResolvedXamlPath) Join-Path (Get-OutputFolder -ResolvedXamlPath $ResolvedXamlPath) 'RoundingEvents.csv' }

    function Load-NearbyRoundingEvents {
        param([string]$ResolvedXamlPath)
        if (-not $script:AppState) { return }
        if (-not ($script:AppState.PSObject.Properties.Name -contains 'NearbyRoundedTodayAssetTags') -or -not $script:AppState.NearbyRoundedTodayAssetTags) {
            $script:AppState | Add-Member -NotePropertyName NearbyRoundedTodayAssetTags -NotePropertyValue (New-Object 'System.Collections.Generic.HashSet[string]') -Force
        }
        $script:AppState.NearbyRoundedTodayAssetTags.Clear()
        $path = Get-RoundingEventsPath -ResolvedXamlPath $ResolvedXamlPath
        if (-not (Test-Path -LiteralPath $path)) { return }
        $today = (Get-Date).Date
        foreach ($row in @(Import-Csv -LiteralPath $path)) {
            $dt = [DateTime]::MinValue
            if (-not [DateTime]::TryParse([string]$row.Timestamp, [ref]$dt) -or $dt.Date -ne $today) { continue }
            $assetKey = [string]$row.AssetTag
            if (-not [string]::IsNullOrWhiteSpace($assetKey)) { [void]$script:AppState.NearbyRoundedTodayAssetTags.Add($assetKey.Trim().ToUpperInvariant()) }
        }
    }

    function Get-RoundingPlanPath { param([string]$ResolvedXamlPath) Join-Path (Get-OutputFolder -ResolvedXamlPath $ResolvedXamlPath) 'RoundingPlan.json' }

    function Read-RoundingPlan {
        param([string]$ResolvedXamlPath)
        $path = Get-RoundingPlanPath -ResolvedXamlPath $ResolvedXamlPath
        if (-not (Test-Path -LiteralPath $path)) { return $null }
        try {
            $plan = Get-Content -LiteralPath $path -Raw | ConvertFrom-Json
            $dates = @($plan.Dates | ForEach-Object { [DateTime]::ParseExact([string]$_, 'yyyy-MM-dd', [Globalization.CultureInfo]::InvariantCulture).Date })
            if ($dates.Count -eq 0) { return $null }
            return [pscustomobject]@{ Dates=($dates | Sort-Object); DailyTarget=30 }
        } catch { return $null }
    }

    function Save-RoundingPlan {
        param([string]$ResolvedXamlPath,[DateTime[]]$Dates)
        Ensure-OutputFolder -ResolvedXamlPath $ResolvedXamlPath
        $path = Get-RoundingPlanPath -ResolvedXamlPath $ResolvedXamlPath
        [pscustomobject]@{ Dates=@($Dates | Sort-Object | ForEach-Object { $_.ToString('yyyy-MM-dd') }); DailyTarget=30 } |
            ConvertTo-Json | Set-Content -LiteralPath $path -Encoding UTF8
    }

    function Show-RoundingPlanDialog {
        param([System.Windows.Window]$Owner,[DateTime[]]$SelectedDates)
        $dialog = New-Object System.Windows.Window
        $dialog.Title = 'Choose rounding days'
        $dialog.Width = 660; $dialog.Height = 440; $dialog.WindowStartupLocation = 'CenterOwner'; $dialog.ResizeMode = 'NoResize'
        $dialog.Background = [System.Windows.Media.BrushConverter]::new().ConvertFromString('#F3F5F7')
        $dialog.Icon = New-CalendarWindowIcon
        if ($Owner) { $dialog.Owner = $Owner }

        $root = New-Object System.Windows.Controls.Grid
        $root.Margin = New-Object System.Windows.Thickness(16)
        $root.RowDefinitions.Add((New-Object System.Windows.Controls.RowDefinition -Property @{ Height='Auto' }))
        $root.RowDefinitions.Add((New-Object System.Windows.Controls.RowDefinition -Property @{ Height='*' }))
        $root.RowDefinitions.Add((New-Object System.Windows.Controls.RowDefinition -Property @{ Height='Auto' }))
        $text = New-Object System.Windows.Controls.TextBlock
        $text.Text = 'Select each day you plan to round devices for this work period. Target is 30 devices per selected day.'
        $text.TextWrapping = 'Wrap'; $text.Margin = New-Object System.Windows.Thickness(0,0,0,14)
        $text.FontSize = 13
        [System.Windows.Controls.Grid]::SetRow($text, 0); $root.Children.Add($text) | Out-Null

        $content = New-Object System.Windows.Controls.Grid
        $content.ColumnDefinitions.Add((New-Object System.Windows.Controls.ColumnDefinition -Property @{ Width='360' }))
        $content.ColumnDefinitions.Add((New-Object System.Windows.Controls.ColumnDefinition -Property @{ Width='18' }))
        $content.ColumnDefinitions.Add((New-Object System.Windows.Controls.ColumnDefinition -Property @{ Width='*' }))
        [System.Windows.Controls.Grid]::SetRow($content, 1); $root.Children.Add($content) | Out-Null

        $calendarHost = New-Object System.Windows.Controls.Border
        $calendarHost.Background = [System.Windows.Media.Brushes]::White
        $calendarHost.BorderBrush = [System.Windows.Media.BrushConverter]::new().ConvertFromString('#D1D8E0')
        $calendarHost.BorderThickness = New-Object System.Windows.Thickness(1)
        $calendarHost.CornerRadius = New-Object System.Windows.CornerRadius(4)
        $calendarHost.Padding = New-Object System.Windows.Thickness(14)
        $calendarHost.HorizontalAlignment = 'Left'; $calendarHost.VerticalAlignment = 'Top'
        $calendar = New-Object System.Windows.Controls.Calendar
        $calendar.SelectionMode = 'MultipleRange'; $calendar.DisplayDate = Get-Date
        $calendar.LayoutTransform = New-Object System.Windows.Media.ScaleTransform(1.42, 1.42)
        foreach ($d in @($SelectedDates)) { if ($d) { $calendar.SelectedDates.Add($d.Date) } }
        $calendarHost.Child = $calendar
        [System.Windows.Controls.Grid]::SetColumn($calendarHost, 0); $content.Children.Add($calendarHost) | Out-Null

        $summary = New-Object System.Windows.Controls.Border
        $summary.Background = [System.Windows.Media.Brushes]::White
        $summary.BorderBrush = [System.Windows.Media.BrushConverter]::new().ConvertFromString('#D1D8E0')
        $summary.BorderThickness = New-Object System.Windows.Thickness(1)
        $summary.CornerRadius = New-Object System.Windows.CornerRadius(4)
        $summary.Padding = New-Object System.Windows.Thickness(16)
        $summary.VerticalAlignment = 'Top'
        $summaryPanel = New-Object System.Windows.Controls.StackPanel
        $summaryTitle = New-Object System.Windows.Controls.TextBlock
        $summaryTitle.Text = 'Adjusting Targets'; $summaryTitle.FontSize = 14; $summaryTitle.FontWeight = 'SemiBold'; $summaryTitle.Margin = New-Object System.Windows.Thickness(0,0,0,14)
        $daysText = New-Object System.Windows.Controls.TextBlock
        $daysText.FontSize = 13; $daysText.Margin = New-Object System.Windows.Thickness(0,0,0,10)
        $devicesText = New-Object System.Windows.Controls.TextBlock
        $devicesText.FontSize = 13; $devicesText.TextWrapping = 'Wrap'
        $summaryPanel.Children.Add($summaryTitle) | Out-Null; $summaryPanel.Children.Add($daysText) | Out-Null; $summaryPanel.Children.Add($devicesText) | Out-Null
        $summary.Child = $summaryPanel
        [System.Windows.Controls.Grid]::SetColumn($summary, 2); $content.Children.Add($summary) | Out-Null

        $updateTargets = { $days = $calendar.SelectedDates.Count; $daysText.Text = "Days selected: $days"; $devicesText.Text = "Devices needed to round: $($days * 30)" }
        $calendar.Add_SelectedDatesChanged($updateTargets); & $updateTargets

        $buttons = New-Object System.Windows.Controls.StackPanel
        $buttons.Orientation = 'Horizontal'; $buttons.HorizontalAlignment = 'Right'; $buttons.Margin = New-Object System.Windows.Thickness(0,14,0,0)
        $ok = New-Object System.Windows.Controls.Button; $ok.Content = 'Save'; $ok.MinWidth = 104; $ok.Height = 32; $ok.Margin = New-Object System.Windows.Thickness(0,0,8,0)
        $ok.Foreground = [System.Windows.Media.Brushes]::White; $ok.Background = [System.Windows.Media.BrushConverter]::new().ConvertFromString('#16A34A'); $ok.BorderBrush = $ok.Background; $ok.FontWeight = 'SemiBold'; $ok.FontSize = 11; $ok.Padding = New-Object System.Windows.Thickness(12,6,12,6); Set-RoundedButtonTemplate -Button $ok -CornerRadius 8
        $cancel = New-Object System.Windows.Controls.Button; $cancel.Content = 'Cancel'; $cancel.MinWidth = 104; $cancel.Height = 32
        $cancel.Foreground = [System.Windows.Media.Brushes]::White; $cancel.Background = [System.Windows.Media.BrushConverter]::new().ConvertFromString('#1F2A44'); $cancel.BorderBrush = $cancel.Background; $cancel.FontWeight = 'SemiBold'; $cancel.FontSize = 11; $cancel.Padding = New-Object System.Windows.Thickness(12,6,12,6); Set-RoundedButtonTemplate -Button $cancel -CornerRadius 8
        $ok.Add_Click({ if ($calendar.SelectedDates.Count -eq 0) { [System.Windows.MessageBox]::Show('Choose at least one rounding day.', 'Rounding days') | Out-Null; return }; $dialog.DialogResult = $true })
        $cancel.Add_Click({ $dialog.DialogResult = $false })
        $buttons.Children.Add($ok) | Out-Null; $buttons.Children.Add($cancel) | Out-Null
        [System.Windows.Controls.Grid]::SetRow($buttons, 2); $root.Children.Add($buttons) | Out-Null
        $dialog.Content = $root
        if ($dialog.ShowDialog()) { return @($calendar.SelectedDates | ForEach-Object { $_.Date } | Sort-Object) }
        return $null
    }

    function Get-RoundingEventCounts {
        param([string]$ResolvedXamlPath,[DateTime[]]$Dates)
        $today = (Get-Date).Date
        $selected = @{}
        foreach ($d in @($Dates)) { $selected[$d.ToString('yyyy-MM-dd')] = $true }
        $todayCount = 0; $weekCount = 0
        $path = Get-RoundingEventsPath -ResolvedXamlPath $ResolvedXamlPath
        if (Test-Path -LiteralPath $path) {
            foreach ($row in @(Import-Csv -LiteralPath $path)) {
                $dt = [DateTime]::MinValue
                if ([DateTime]::TryParse([string]$row.Timestamp, [ref]$dt)) {
                    $key = $dt.Date.ToString('yyyy-MM-dd')
                    if ($selected.ContainsKey($key)) { $weekCount++ }
                    if ($dt.Date -eq $today) { $todayCount++ }
                }
            }
        }
        return [pscustomobject]@{ Today=$todayCount; Week=$weekCount }
    }

    function Get-RemainingRoundingDays {
        param([DateTime[]]$Dates,[DateTime]$Today)
        return @($Dates | Where-Object { $_.Date -ge $Today.Date }).Count
    }

    function Get-TodayRoundingTarget {
        param([int]$WeekRemaining,[int]$RemainingDays)
        if ($WeekRemaining -le 0 -or $RemainingDays -le 0) { return 0 }
        return [int][Math]::Ceiling($WeekRemaining / [double]$RemainingDays)
    }

    function Update-RoundingPlanBadges {
        param([hashtable]$Ui,[string]$ResolvedXamlPath,[pscustomobject]$Plan)
        $days = @($Plan.Dates)
        $today = (Get-Date).Date
        $dailyTarget = 30
        $weekTarget = $dailyTarget * $days.Count
        $counts = Get-RoundingEventCounts -ResolvedXamlPath $ResolvedXamlPath -Dates $days
        $weekRemaining = [Math]::Max(0, $weekTarget - $counts.Week)
        $remainingDays = Get-RemainingRoundingDays -Dates $days -Today $today
        $todayTarget = Get-TodayRoundingTarget -WeekRemaining $weekRemaining -RemainingDays $remainingDays
        Set-ControlText -Control $Ui.DaysPerWeekBadgeText -Value "Days/week $($days.Count)"
        Set-ControlText -Control $Ui.TodayBadgeText -Value "Today $($counts.Today) / $todayTarget"
        Set-ControlText -Control $Ui.ThisWeekBadgeText -Value "This week $($counts.Week) / $weekTarget"
        Set-ControlText -Control $Ui.RemainingPerDayBadgeText -Value "Week remaining $weekRemaining"

        if ($counts.Today -gt 75) { Set-BadgeStyle -Border $Ui.TodayBadge -BackgroundHex '#FCE3E5' -ForegroundHex '#BE123C' }
        else { Set-BadgeStyle -Border $Ui.TodayBadge -BackgroundHex '#FEE7C3' -ForegroundHex '#B45309' }

        if ($counts.Week -ge 150) { Set-BadgeStyle -Border $Ui.ThisWeekBadge -BackgroundHex '#DDF7E5' -ForegroundHex '#15803D' }
        else { Set-BadgeStyle -Border $Ui.ThisWeekBadge -BackgroundHex '#FEE7C3' -ForegroundHex '#B45309' }

        if ($todayTarget -le 35) { Set-BadgeStyle -Border $Ui.RemainingPerDayBadge -BackgroundHex '#DDF7E5' -ForegroundHex '#15803D' }
        elseif ($todayTarget -le 75) { Set-BadgeStyle -Border $Ui.RemainingPerDayBadge -BackgroundHex '#FEE7C3' -ForegroundHex '#B45309' }
        else { Set-BadgeStyle -Border $Ui.RemainingPerDayBadge -BackgroundHex '#FCE3E5' -ForegroundHex '#BE123C' }
    }

    function Ensure-RoundingPlan {
        param([hashtable]$Ui,[System.Windows.Window]$Window,[string]$ResolvedXamlPath,[bool]$Force)
        $plan = Read-RoundingPlan -ResolvedXamlPath $ResolvedXamlPath
        $today = (Get-Date).Date
        $needsPrompt = $Force -or -not $plan -or $today -lt @($plan.Dates)[0].Date -or $today -gt @($plan.Dates)[@($plan.Dates).Count - 1].Date
        if ($needsPrompt) {
            $existing = if ($plan) { @($plan.Dates) } else { @($today) }
            $chosen = Show-RoundingPlanDialog -Owner $Window -SelectedDates $existing
            if ($chosen) { Save-RoundingPlan -ResolvedXamlPath $ResolvedXamlPath -Dates $chosen; $plan = Read-RoundingPlan -ResolvedXamlPath $ResolvedXamlPath }
        }
        if (-not $plan) { $plan = [pscustomobject]@{ Dates=@($today); DailyTarget=30 } }
        $script:RoundingPlan = $plan
        Update-RoundingPlanBadges -Ui $Ui -ResolvedXamlPath $ResolvedXamlPath -Plan $plan
    }
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

    function Show-CopiedClipboardToast {
        param([object]$Control)
        if (-not $Control) { return }

        $toastText = New-Object System.Windows.Controls.TextBlock
        $toastText.Text = 'Copied to clipboard'
        $toastText.Foreground = [System.Windows.Media.Brushes]::White
        $toastText.FontSize = 12
        $toastText.FontWeight = [System.Windows.FontWeights]::SemiBold
        $toastText.Margin = New-Object System.Windows.Thickness(10,5,10,5)
        $toastText.IsHitTestVisible = $false

        $toastBorder = New-Object System.Windows.Controls.Border
        $toastBorder.Background = New-Object System.Windows.Media.SolidColorBrush([System.Windows.Media.Color]::FromArgb(235, 31, 42, 68))
        $toastBorder.CornerRadius = New-Object System.Windows.CornerRadius(4)
        $toastBorder.Child = $toastText
        $toastBorder.IsHitTestVisible = $false

        $popup = New-Object System.Windows.Controls.Primitives.Popup
        $popup.AllowsTransparency = $true
        $popup.Child = $toastBorder
        $popup.HorizontalOffset = 14
        $popup.VerticalOffset = 16
        $popup.Placement = [System.Windows.Controls.Primitives.PlacementMode]::MousePoint
        $popup.PlacementTarget = $Control
        $popup.IsOpen = $true

        $timer = New-Object System.Windows.Threading.DispatcherTimer
        $timer.Interval = [TimeSpan]::FromMilliseconds(1500)
        $timer.Tag = $popup
        $timer.Add_Tick({
            param($sender,$eventArgs)
            $sender.Stop()
            if ($sender.Tag) { $sender.Tag.IsOpen = $false }
        })
        $timer.Start()
    }

    function Copy-SummaryValueToClipboard {
        param([hashtable]$Ui,[object]$Control,[string]$FieldName)
        if (-not $Control) { return }
        $value = ''
        try { $value = [string]$Control.Text } catch {}
        if ([string]::IsNullOrWhiteSpace($value)) { return }
        try {
            [System.Windows.Clipboard]::SetText($value.Trim())
            Show-CopiedClipboardToast -Control $Control
            Set-StatusMessage -Ui $Ui -Mode 'Found' -CustomText "Copied $FieldName to clipboard."
        } catch {
            Set-StatusMessage -Ui $Ui -Mode 'Warning' -CustomText "Unable to copy $FieldName to clipboard."
        }
    }

    function Register-SummaryClipboardCopy {
        param([hashtable]$Ui)
        $summaryCopyFields = @(
            @{ Control = $Ui.HostNameDisplay; FieldName = 'Name' },
            @{ Control = $Ui.AssetTagDisplay; FieldName = 'Asset Tag' },
            @{ Control = $Ui.SerialDisplay; FieldName = 'Serial' },
            @{ Control = $Ui.ParentDisplay; FieldName = 'Parent' },
            @{ Control = $Ui.RitmDisplay; FieldName = 'PO/RITM' }
        )
        foreach ($field in $summaryCopyFields) {
            $control = $field.Control
            $fieldName = $field.FieldName
            if (-not $control) { continue }
            $control.Cursor = [System.Windows.Input.Cursors]::Hand
            $control.ToolTip = "Double-click to copy $fieldName to the clipboard."
            $control.Add_MouseLeftButtonDown({
                param($sender,$eventArgs)
                if ($eventArgs.ClickCount -eq 2) {
                    Copy-SummaryValueToClipboard -Ui $sender.Tag.Ui -Control $sender -FieldName $sender.Tag.FieldName
                    $eventArgs.Handled = $true
                }
            })
            $control.Tag = [pscustomobject]@{ Ui = $Ui; FieldName = $fieldName }
        }
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
        foreach ($name in @($Names)) {
            if ([string]::IsNullOrWhiteSpace($name)) { continue }
            foreach ($property in @($Row.PSObject.Properties)) {
                if ($property.Name -ne $name) { continue }
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



    function Get-DataFileInfo {
        param([string]$ResolvedXamlPath,[string]$SiteFolderPath)
        $dataRoot = Join-Path (Split-Path -Parent $ResolvedXamlPath) 'Data'
        $useSiteFolder = -not [string]::IsNullOrWhiteSpace($SiteFolderPath) -and (Test-Path -LiteralPath $SiteFolderPath)
        $sitePath = if ($useSiteFolder) { $SiteFolderPath } else { $dataRoot }
        $definitions = @(
            @{ Label='Computers'; Path=$sitePath; Filter='Computers - *.csv'; Recurse=(-not $useSiteFolder) },
            @{ Label='Monitors'; Path=$sitePath; Filter='Monitors - *.csv'; Recurse=(-not $useSiteFolder) },
            @{ Label='Mics'; Path=$dataRoot; Filter='Mics.csv' },
            @{ Label='Scanners'; Path=$dataRoot; Filter='Scanners.csv' },
            @{ Label='Carts'; Path=$dataRoot; Filter='Carts.csv' },
            @{ Label='Rounding'; Path=$dataRoot; Filter='Rounding.csv' }
        )
        $files = @()
        foreach ($definition in $definitions) {
            if (-not (Test-Path -LiteralPath $definition.Path)) {
                $files += [pscustomobject]@{ Label=$definition.Label; Name=$definition.Filter; Path=''; LastWriteTime=$null; Age=$null; Missing=$true }
                continue
            }
            $getChildItemParams = @{ LiteralPath=$definition.Path; File=$true; Filter=$definition.Filter; ErrorAction='SilentlyContinue' }
            if ($definition.ContainsKey('Recurse') -and $definition.Recurse) { $getChildItemParams.Recurse = $true }
            $matches = @(Get-ChildItem @getChildItemParams | Sort-Object FullName)
            if ($matches.Count -eq 0) {
                $files += [pscustomobject]@{ Label=$definition.Label; Name=$definition.Filter; Path=''; LastWriteTime=$null; Age=$null; Missing=$true }
                continue
            }
            foreach ($file in $matches) {
                $age = (Get-Date) - $file.LastWriteTime
                $files += [pscustomobject]@{ Label=$definition.Label; Name=$file.Name; Path=$file.FullName; LastWriteTime=$file.LastWriteTime; Age=$age; Missing=$false }
            }
        }
        return @($files)
    }

    function Format-DataFileAge {
        param([TimeSpan]$Age)
        if ($null -eq $Age) { return 'missing' }

        $units = @(
            @{ Name='year'; Value=($Age.TotalDays / 365) },
            @{ Name='week'; Value=($Age.TotalDays / 7) },
            @{ Name='day'; Value=$Age.TotalDays },
            @{ Name='hour'; Value=$Age.TotalHours }
        )
        $unit = $units | Where-Object { $_.Value -ge 1 } | Select-Object -First 1
        if ($null -eq $unit) { $unit = @{ Name='hour'; Value=0 } }

        $value = [Math]::Round([double]$unit.Value, 1)
        $displayValue = if ($value -eq [Math]::Truncate($value)) { [int]$value } else { $value.ToString('0.0') }
        $plural = if ($value -eq 1) { '' } else { 's' }
        return "$displayValue $($unit.Name)$plural old"
    }

    function Add-DataFileInfoLine {
        param(
            [System.Windows.Controls.Panel]$Panel,
            [string]$Label,
            [string]$Name,
            [string]$StatusText,
            [System.Windows.Media.Brush]$StatusBrush
        )

        $textBlock = New-Object System.Windows.Controls.TextBlock
        $textBlock.FontSize = 13
        $textBlock.Margin = New-Object System.Windows.Thickness(0,0,0,6)
        $textBlock.TextWrapping = 'Wrap'

        $labelRun = New-Object System.Windows.Documents.Run("$Label`n")
        $labelRun.FontWeight = [System.Windows.FontWeights]::Bold
        $labelRun.Foreground = [System.Windows.Media.BrushConverter]::new().ConvertFromString('#1F2937')
        $null = $textBlock.Inlines.Add($labelRun)
        $null = $textBlock.Inlines.Add((New-Object System.Windows.Documents.Run("$Name - ")))
        $statusRun = New-Object System.Windows.Documents.Run($StatusText)
        $statusRun.Foreground = $StatusBrush
        $null = $textBlock.Inlines.Add($statusRun)
        $null = $Panel.Children.Add($textBlock)
    }

    function Show-DataFileInfo {
        param([hashtable]$Ui)
        $files = @($script:AppState.DataFiles)
        if ($files.Count -eq 0) { [System.Windows.MessageBox]::Show('No data files were found.', 'Data File') | Out-Null; return }

        $dialog = New-Object System.Windows.Window
        $dialog.Title = 'Data File Ages'
        $dialog.SizeToContent = 'WidthAndHeight'
        $dialog.MinWidth = 360
        $dialog.MaxWidth = 560
        $dialog.WindowStartupLocation = 'CenterOwner'
        $dialog.ResizeMode = 'NoResize'
        $dialog.Background = [System.Windows.Media.Brushes]::White
        if ($Ui.ContainsKey('Window') -and $Ui.Window) { $dialog.Owner = $Ui.Window }

        $panel = New-Object System.Windows.Controls.StackPanel
        $panel.Margin = New-Object System.Windows.Thickness(18)

        $title = New-Object System.Windows.Controls.TextBlock
        $title.Text = 'Data File Ages'
        $title.FontSize = 16
        $title.FontWeight = [System.Windows.FontWeights]::SemiBold
        $title.Margin = New-Object System.Windows.Thickness(0,0,0,12)
        $null = $panel.Children.Add($title)

        foreach ($file in $files) {
            if ($file.Missing) {
                Add-DataFileInfoLine -Panel $panel -Label $file.Label -Name $file.Name -StatusText 'missing' -StatusBrush ([System.Windows.Media.BrushConverter]::new().ConvertFromString('#BE123C'))
            }
            else {
                Add-DataFileInfoLine -Panel $panel -Label $file.Label -Name $file.Name -StatusText (Format-DataFileAge -Age $file.Age) -StatusBrush ([System.Windows.Media.BrushConverter]::new().ConvertFromString('#374151'))
            }
        }

        $ok = New-Object System.Windows.Controls.Button
        $ok.Content = 'OK'
        $ok.Width = 78
        $ok.Margin = New-Object System.Windows.Thickness(0,10,0,0)
        $ok.HorizontalAlignment = 'Right'
        $ok.IsDefault = $true
        $ok.Add_Click({ $dialog.Close() })
        $null = $panel.Children.Add($ok)

        $dialog.Content = $panel
        $dialog.ShowDialog() | Out-Null
    }

    function Update-DataFileBadge {
        param([hashtable]$Ui,[object[]]$DataFiles)
        Set-BadgeText -Border $Ui.DataFileBadge -Text 'Data File'
        $computerFiles = @($DataFiles | Where-Object { $_.Label -eq 'Computers' -and -not $_.Missing })
        if ($computerFiles.Count -eq 0) {
            Set-BadgeStyle -Border $Ui.DataFileBadge -BackgroundHex '#FCE3E5' -ForegroundHex '#BE123C'
            return
        }
        $computerAgeDays = (($computerFiles | Sort-Object LastWriteTime -Descending | Select-Object -First 1).Age).TotalDays
        if ($computerAgeDays -lt 1) { Set-BadgeStyle -Border $Ui.DataFileBadge -BackgroundHex '#DDF7E5' -ForegroundHex '#15803D' }
        elseif ($computerAgeDays -le 3) { Set-BadgeStyle -Border $Ui.DataFileBadge -BackgroundHex '#FEE7C3' -ForegroundHex '#B45309' }
        else { Set-BadgeStyle -Border $Ui.DataFileBadge -BackgroundHex '#FCE3E5' -ForegroundHex '#BE123C' }
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
            Model=Get-FieldValue -Row $Row -Names @('model_id','Model')
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
        $results = @([pscustomobject]@{ Role='Parent'; Type=$parentType; Name=$effectiveParent.Name; AssetTag=$effectiveParent.AssetTag; Serial=$effectiveParent.Serial; Model=$effectiveParent.Model; SerialForeground='#1F2937'; SerialToolTip=''; RITM=$effectiveParent.RITM; Retire=(Format-DateLong $effectiveParent.RetireDate); CmdbUrl=(Get-CmdbLink -DeviceType $effectiveParent.DetectedType -AssetTag $effectiveParent.AssetTag); Device=$effectiveParent })
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
                $record = [pscustomobject]@{ Role=$role; Type=$type; Name=$childDevice.Name; AssetTag=$childDevice.AssetTag; Serial=$childDevice.Serial; Model=$childDevice.Model; SerialForeground='#1F2937'; SerialToolTip=''; RITM=$childDevice.RITM; Retire=(Format-DateLong $childDevice.RetireDate); CmdbUrl=(Get-CmdbLink -DeviceType $type -AssetTag $childDevice.AssetTag); Device=$childDevice }
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

    function Invoke-PingOnce {
        param([string]$ComputerName,[int]$TimeoutMilliseconds=1000)
        if ([string]::IsNullOrWhiteSpace($ComputerName)) { return $null }
        $ping = $null
        try {
            $ping = New-Object System.Net.NetworkInformation.Ping
            return $ping.Send($ComputerName, $TimeoutMilliseconds)
        } catch {
            return $null
        } finally {
            try { if ($ping) { $ping.Dispose() } } catch {}
        }
    }

    function Test-ComputerPingable {
        param([string]$ComputerName)
        $reply = Invoke-PingOnce -ComputerName $ComputerName -TimeoutMilliseconds 1000
        return ($reply -and $reply.Status -eq [System.Net.NetworkInformation.IPStatus]::Success)
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

        $reply = Invoke-PingOnce -ComputerName $ComputerName -TimeoutMilliseconds 1000
        if (-not $reply -or $reply.Status -ne [System.Net.NetworkInformation.IPStatus]::Success) {
            $errorMessage = if ($reply) { [string]$reply.Status } else { 'Timed out' }
            return [pscustomobject]@{ HostName=$ComputerName; Success=$false; IpAddress='Unknown'; ResponseTime='Timed out'; Subnet='Unknown'; ErrorMessage=$errorMessage }
        }

        $ipAddress = Get-IPv4AddressFromPingReply -Reply $reply -ComputerName $ComputerName
        if ([string]::IsNullOrWhiteSpace($ipAddress)) { $ipAddress = 'Unknown' }

        $responseMs = Get-FieldValue -Row $reply -Names @('RoundtripTime','ResponseTime','Latency')
        if ([string]::IsNullOrWhiteSpace($responseMs)) { $responseMs = 'Unknown' } else { $responseMs = "$responseMs ms" }

        return [pscustomobject]@{ HostName=$ComputerName; Success=$true; IpAddress=$ipAddress; ResponseTime=$responseMs; Subnet=(Resolve-SubnetName -IpAddress $ipAddress -DataRoot $DataRoot); ErrorMessage='' }
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
        $result = [pscustomobject]@{ ComputerSerial=$null; MonitorSerials=@(); MonitorDetails=@(); Offline=$false }
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
                $name = Convert-WmiIdToString -Id $m.UserFriendlyName
                $manufacturer = Convert-WmiIdToString -Id $m.ManufacturerName
                $productCode = Convert-WmiIdToString -Id $m.ProductCodeID
                $serialText = if ([string]::IsNullOrWhiteSpace($serial)) { '' } else { $serial.Trim() }
                if (-not [string]::IsNullOrWhiteSpace($serialText)) { $result.MonitorSerials += $serialText }
                $detail = [pscustomobject]@{
                    Type='Monitor'; Name=$name; Manufacturer=$manufacturer; ProductCode=$productCode; Serial=$serialText
                    ManufactureWeek=$m.WeekOfManufacture; ManufactureYear=$m.YearOfManufacture; InstanceName=$m.InstanceName; Active=$m.Active
                }
                $result.MonitorDetails += $detail
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

    function Get-ConnectedMonitorDetailBySerial {
        param([object[]]$MonitorDetails,[string]$Serial)
        if ([string]::IsNullOrWhiteSpace($Serial)) { return $null }
        $target = $Serial.Trim().ToUpper()
        $compactTarget = ($target -replace '[-\s]','')
        foreach ($detail in @($MonitorDetails)) {
            if (-not $detail -or [string]::IsNullOrWhiteSpace($detail.Serial)) { continue }
            $detailSerial = $detail.Serial.Trim().ToUpper()
            $compactDetailSerial = ($detailSerial -replace '[-\s]','')
            if ($detailSerial -eq $target -or $compactDetailSerial -eq $compactTarget) { return $detail }
        }
        return $null
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
            $changed = Show-AddPeripheralDialog -Ui $Ui -ParentDevice $ParentDevice -Inventory $Inventory -ResolvedXamlPath $ResolvedXamlPath -DefaultSearchText $serialToLink -InfoMessage 'This monitor is connected but not linked. Click Add to link if it is found in inventory.' -ConnectedMonitorDetails @($wmiData.MonitorDetails)
            if ($changed) { $Ui.AssociatedDevicesDataGrid.ItemsSource = Build-AssociatedDevices -Device $ParentDevice -Inventory $Inventory }
        }
    }

    function Show-AddPeripheralDialog {
        param([hashtable]$Ui,[pscustomobject]$ParentDevice,[pscustomobject]$Inventory,[string]$ResolvedXamlPath,[string]$DefaultSearchText='',[string]$InfoMessage='',[object[]]$ConnectedMonitorDetails=@())
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
        & $addPreviewRow 7 'Manufacturer:' 'Manufacturer'
        & $addPreviewRow 8 'Product Code:' 'ProductCode'
        & $addPreviewRow 9 'Manufactured:' 'Manufactured'
        & $addPreviewRow 10 'WMI Instance:' 'InstanceName'
        & $addPreviewRow 11 'Active:' 'Active'
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
                $previewValues.Type.Text = if ($result -and $result.NormalizedInput) { '(not found in inventory)' } else { '' }
                $detail = $null
                if ($result -and $result.NormalizedInput) { $detail = Get-ConnectedMonitorDetailBySerial -MonitorDetails $ConnectedMonitorDetails -Serial $result.NormalizedInput }
                if (-not $detail) { $detail = Get-ConnectedMonitorDetailBySerial -MonitorDetails $ConnectedMonitorDetails -Serial $txt.Text }
                if ($detail) {
                    $previewValues.Type.Text = 'Connected monitor (not in inventory)'
                    $previewValues.Name.Text = $detail.Name
                    $previewValues.Serial.Text = $detail.Serial
                    $previewValues.Manufacturer.Text = $detail.Manufacturer
                    $previewValues.ProductCode.Text = $detail.ProductCode
                    $previewValues.Manufactured.Text = if ($detail.ManufactureYear) { "Week $($detail.ManufactureWeek), $($detail.ManufactureYear)" } else { '' }
                    $previewValues.InstanceName.Text = $detail.InstanceName
                    $previewValues.Active.Text = $detail.Active
                }
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
            $detail = Get-ConnectedMonitorDetailBySerial -MonitorDetails $ConnectedMonitorDetails -Serial $c.Serial
            if ($detail) {
                $previewValues.Manufacturer.Text = $detail.Manufacturer
                $previewValues.ProductCode.Text = $detail.ProductCode
                $previewValues.Manufactured.Text = if ($detail.ManufactureYear) { "Week $($detail.ManufactureWeek), $($detail.ManufactureYear)" } else { '' }
                $previewValues.InstanceName.Text = $detail.InstanceName
                $previewValues.Active.Text = $detail.Active
            }
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

    function Get-NearbyScopeKey {
        param([string]$Location)
        $normalized = Normalize-LocationValue -Value $Location
        if ([string]::IsNullOrWhiteSpace($normalized)) { return $null }
        return $normalized
    }

    function Ensure-NearbyState {
        if (-not $script:AppState) { return }
        if (-not ($script:AppState.PSObject.Properties.Name -contains 'ActiveNearbyScopes') -or -not $script:AppState.ActiveNearbyScopes) {
            $script:AppState | Add-Member -NotePropertyName ActiveNearbyScopes -NotePropertyValue (New-Object 'System.Collections.Generic.HashSet[string]') -Force
        }
    }

    function Add-NearbyScope {
        param([pscustomobject]$Device)
        if (-not $Device) { return $false }
        Ensure-NearbyState
        if (-not $script:AppState -or -not $script:AppState.ActiveNearbyScopes) { return $false }
        $key = Get-NearbyScopeKey -Location $Device.Location
        if (-not $key) { return $false }
        return $script:AppState.ActiveNearbyScopes.Add($key)
    }

    function Test-DeviceInNearbyScope {
        param([pscustomobject]$Device)
        if (-not $Device) { return $false }
        Ensure-NearbyState
        if (-not $script:AppState -or -not $script:AppState.ActiveNearbyScopes -or $script:AppState.ActiveNearbyScopes.Count -eq 0) { return $false }
        $key = Get-NearbyScopeKey -Location $Device.Location
        if (-not $key) { return $false }
        return $script:AppState.ActiveNearbyScopes.Contains($key)
    }

    function Build-NearbyDevices {
        param([pscustomobject]$Device,[pscustomobject]$Inventory)
        if (-not $Inventory -or -not $Inventory.Computers) { return @() }
        Ensure-NearbyState
        $rows = @()
        $seenAssetTags = @{}
        foreach ($row in @($Inventory.Computers)) {
            $computer = ConvertTo-DeviceRecord -Row $row -DetectedType 'Computer'
            if (-not (Test-DeviceInNearbyScope -Device $computer)) { continue }
            $assetKey = if ($computer.AssetTag) { $computer.AssetTag.Trim().ToUpperInvariant() } else { '' }
            if (-not [string]::IsNullOrWhiteSpace($assetKey)) {
                if ($seenAssetTags.ContainsKey($assetKey)) { continue }
                $seenAssetTags[$assetKey] = $true
            }
            $lastRoundedDate = Parse-DateLoose -Value $computer.LastRounded
            $daysAgo = ''
            if ($lastRoundedDate) {
                $daysAgo = [int]((Get-Date).Date - $lastRoundedDate.Date).TotalDays
            }
            $maintenanceType = Get-FieldValue -Row $row -Names @('u_device_rounding','MaintenanceType')
            $rows += [pscustomobject]@{
                HostName=$computer.Name
                IPAddress=''
                Subnet=''
                AssetTag=$computer.AssetTag
                Location=$computer.Location
                Building=$computer.Building
                Floor=$computer.Floor
                Room=$computer.Room
                Department=$computer.Department
                MaintenanceType=(Get-MaintenanceTypeOrDefault -MaintenanceType $maintenanceType -DeviceName $computer.Name)
                LastRounded=(Format-DateLong $computer.LastRounded)
                DaysAgo=$daysAgo
                Status='-'
                StatusOptions=@('-','Inaccessible - Asset not found','Inaccessible - In storage','Inaccessible - In use by Customer','Inaccessible - Laptop is not onsite','Inaccessible - Other','Inaccessible - Restricted area','Inaccessible - Room locked - Card Swipe','Inaccessible - Room locked - Key Lock','Inaccessible - Under renovation','Inaccessible - User working at home')
                IsStatusEditable=$true
                Device=$computer
            }
        }
        return @($rows | Sort-Object Location,Building,Floor,Room,HostName)
    }

    function Update-NearbySummary {
        param([hashtable]$Ui)
        Ensure-NearbyState
        $scopeCount = if ($script:AppState -and $script:AppState.ActiveNearbyScopes) { $script:AppState.ActiveNearbyScopes.Count } else { 0 }
        $rowCount = 0
        try { $rowCount = @($Ui.NearbyDataGrid.ItemsSource).Count } catch {}
        $text = "Nearby scopes (Location): $scopeCount"
        if ($rowCount -gt 0) { $text += " - Showing $rowCount" }
        Set-ControlText -Control $Ui.NearbyScopeSummaryText -Value $text
    }

    function Update-NearbyRows {
        param([hashtable]$Ui,[pscustomobject]$Inventory,[string]$ResolvedXamlPath='')
        $rows = Build-NearbyDevices -Device $script:AppState.CurrentDevice -Inventory $Inventory
        $Ui.NearbyDataGrid.ItemsSource = $rows
        if ($script:AppState) {
            $associated = @()
            try { $associated = @($Ui.AssociatedDevicesDataGrid.ItemsSource) } catch {}
            $script:AppState.SampleData = [pscustomobject]@{ Device=$script:AppState.CurrentDevice; Associated=$associated; Nearby=$rows }
        }
        Update-NearbySummary -Ui $Ui
    }



    function Get-NearbySelectedRows {
        param([hashtable]$Ui)
        if (-not $Ui -or -not $Ui.NearbyDataGrid) { return @() }
        $selected = @($Ui.NearbyDataGrid.SelectedItems | Where-Object { $_ })
        if ($selected.Count -gt 0) { return $selected }
        if ($Ui.NearbyDataGrid.SelectedItem) { return @($Ui.NearbyDataGrid.SelectedItem) }
        return @()
    }

    function Set-NearbySelectedStatus {
        param([hashtable]$Ui,[string]$Status)
        $selected = @(Get-NearbySelectedRows -Ui $Ui)
        if ($selected.Count -eq 0) {
            [System.Windows.MessageBox]::Show('Select one or more Nearby devices first.', 'Nearby Status') | Out-Null
            return
        }
        foreach ($row in $selected) {
            if (-not $row) { continue }
            if ($row.PSObject.Properties.Name -contains 'IsStatusEditable' -and -not [bool]$row.IsStatusEditable) { continue }
            $row.Status = $Status
        }
        try { $Ui.NearbyDataGrid.Items.Refresh() } catch {}
        Set-ControlText -Control $Ui.NearbyScopeSummaryText -Value "Updated status for $($selected.Count) selected Nearby device(s)."
    }

    function Invoke-SelectedNearbyPing {
        param([hashtable]$Ui,[string]$DataRoot)
        $selected = @(Get-NearbySelectedRows -Ui $Ui)
        if ($selected.Count -eq 0) {
            [System.Windows.MessageBox]::Show('Select one or more Nearby devices first.', 'Ping selected host(s)') | Out-Null
            return
        }
        $updated = 0
        foreach ($row in $selected) {
            if (-not $row -or [string]::IsNullOrWhiteSpace([string]$row.HostName)) { continue }
            $result = Invoke-DevicePing -ComputerName ([string]$row.HostName) -DataRoot $DataRoot
            $row.IPAddress = if ($result.IpAddress) { [string]$result.IpAddress } else { 'Unknown' }
            $row.Subnet = if ($result.Subnet) { [string]$result.Subnet } else { 'Unknown' }
            $updated++
        }
        try { $Ui.NearbyDataGrid.Items.Refresh() } catch {}
        Set-ControlText -Control $Ui.NearbyScopeSummaryText -Value "Ping updated $updated selected Nearby host(s)."
    }

    function Initialize-NearbyContextMenu {
        param([hashtable]$Ui,[string]$DataRoot)
        if (-not $Ui -or -not $Ui.NearbyDataGrid) { return }
        $menu = New-Object System.Windows.Controls.ContextMenu
        $pingItem = New-Object System.Windows.Controls.MenuItem -Property @{ Header='Ping selected host(s)' }
        $pingItem.Add_Click({ param($sender,$e) Invoke-SelectedNearbyPing -Ui $ui -DataRoot $dataRoot })
        [void]$menu.Items.Add($pingItem)
        [void]$menu.Items.Add((New-Object System.Windows.Controls.Separator))
        foreach ($status in @('-','Inaccessible - Asset not found','Inaccessible - In storage','Inaccessible - In use by Customer','Inaccessible - Laptop is not onsite','Inaccessible - Other','Inaccessible - Restricted area','Inaccessible - Room locked - Card Swipe','Inaccessible - Room locked - Key Lock','Inaccessible - Under renovation','Inaccessible - User working at home')) {
            $item = New-Object System.Windows.Controls.MenuItem -Property @{ Header=$status; Tag=$status }
            $item.Add_Click({ param($sender,$e) Set-NearbySelectedStatus -Ui $ui -Status ([string]$sender.Tag) })
            [void]$menu.Items.Add($item)
        }
        $Ui.NearbyDataGrid.ContextMenu = $menu
        $Ui.NearbyDataGrid.Add_PreviewMouseRightButtonDown({
            if (-not $this.SelectedItem -and $this.CurrentItem) { $this.SelectedItem = $this.CurrentItem }
        })
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

    function Get-LocationHierarchyFieldValue {
        param([object]$Row,[string]$Name)
        return Get-LocationFieldValue -Row $Row -Name $Name
    }

    function New-LocationHierarchyRow {
        param([string]$City,[string]$Location,[string]$Building,[string]$Floor,[string]$Room,[string]$Department)
        return [pscustomobject]@{ City=[string]$City; Location=[string]$Location; Building=[string]$Building; Floor=[string]$Floor; Room=[string]$Room; Department=[string]$Department }
    }

    function Add-LocationHierarchyRow {
        param([System.Collections.Generic.List[object]]$Rows,[hashtable]$Seen,[object]$Row)
        if (-not $Rows -or -not $Seen -or -not $Row) { return }
        $city = Get-LocationFieldValue $Row 'City'
        $location = Get-LocationFieldValue $Row 'Location'
        if ([string]::IsNullOrWhiteSpace($location)) { $location = Get-FieldValue -Row $Row -Names @('location') }
        $building = Get-LocationFieldValue $Row 'Building'
        if ([string]::IsNullOrWhiteSpace($building)) { $building = Get-FieldValue -Row $Row -Names @('u_building') }
        $floor = Get-LocationFieldValue $Row 'Floor'
        if ([string]::IsNullOrWhiteSpace($floor)) { $floor = Get-FieldValue -Row $Row -Names @('u_floor') }
        $room = Get-LocationFieldValue $Row 'Room'
        if ([string]::IsNullOrWhiteSpace($room)) { $room = Get-FieldValue -Row $Row -Names @('u_room') }
        $department = Get-LocationFieldValue $Row 'Department'
        if ([string]::IsNullOrWhiteSpace($department)) { $department = Get-FieldValue -Row $Row -Names @('u_department_location') }
        $key = '{0}|{1}|{2}|{3}|{4}|{5}' -f (Normalize-LocationValue $city),(Normalize-LocationValue $location),(Normalize-LocationValue $building),(Normalize-LocationValue $floor),(Normalize-LocationValue $room),(Normalize-LocationValue $department)
        if ($Seen.ContainsKey($key)) { return }
        $Seen[$key] = $true
        [void]$Rows.Add((New-LocationHierarchyRow -City $city -Location $location -Building $building -Floor $floor -Room $room -Department $department))
    }

    function Get-LocationHierarchyRows {
        param([pscustomobject]$Inventory)
        $rows = New-Object System.Collections.Generic.List[object]
        $seen = @{}
        if (-not $Inventory) { return @() }
        foreach ($row in @($Inventory.Locations)) { Add-LocationHierarchyRow -Rows $rows -Seen $seen -Row $row }
        foreach ($row in @($Inventory.Computers)) { Add-LocationHierarchyRow -Rows $rows -Seen $seen -Row $row }
        return @($rows.ToArray())
    }

    function Test-LocationColumnValue {
        param([pscustomobject]$Inventory,[string]$Value,[string]$Column)
        if ([string]::IsNullOrWhiteSpace($Value)) { return $false }
        $n = Normalize-LocationValue $Value
        foreach ($row in @(Get-LocationHierarchyRows -Inventory $Inventory)) {
            if ((Normalize-LocationValue $row.$Column) -eq $n) { return $true }
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
        foreach ($row in @(Get-LocationHierarchyRows -Inventory $Inventory)) {
            $room = $row.Room
            if ((Normalize-LocationValue $room) -eq $n) { return $true }
            if ($code -and (Extract-RoomCode $room) -eq $code) { return $true }
        }
        return $false
    }

    function Set-LocationValidationStyle {
        param([hashtable]$Ui,[pscustomobject]$Inventory)
        $validBrush = New-Brush '#CCF2D3'; $validBorder = New-Brush '#7CE0A6'
        $badBrush = New-Brush '#FCE3E5'; $badBorder = New-Brush '#F5A3AA'
        $rows = @(Get-LocationHierarchyRows -Inventory $Inventory)
        if ($rows.Count -eq 0) {
            foreach ($box in @($Ui.CityTextBox,$Ui.LocationTextBox,$Ui.BuildingTextBox,$Ui.FloorTextBox,$Ui.RoomTextBox,$Ui.DepartmentTextBox)) {
                if (-not $box) { continue }
                $box.Background = $validBrush
                $box.BorderBrush = $validBorder
            }
            return
        }
        $checks = @(
            [pscustomobject]@{ Box=$Ui.CityTextBox;       IsValid=(Test-LocationColumnValue $Inventory $Ui.CityTextBox.Text 'City') },
            [pscustomobject]@{ Box=$Ui.LocationTextBox;   IsValid=(Test-LocationColumnValue $Inventory $Ui.LocationTextBox.Text 'Location') },
            [pscustomobject]@{ Box=$Ui.BuildingTextBox;   IsValid=(Test-LocationColumnValue $Inventory $Ui.BuildingTextBox.Text 'Building') },
            [pscustomobject]@{ Box=$Ui.FloorTextBox;      IsValid=(Test-LocationColumnValue $Inventory $Ui.FloorTextBox.Text 'Floor') },
            [pscustomobject]@{ Box=$Ui.RoomTextBox;       IsValid=(Test-LocationRoomValue $Inventory $Ui.RoomTextBox.Text) },
            [pscustomobject]@{ Box=$Ui.DepartmentTextBox; IsValid=(Test-LocationColumnValue $Inventory $Ui.DepartmentTextBox.Text 'Department') }
        )
        foreach ($item in $checks) {
            if (-not $item.Box) { continue }
            $item.Box.Background = if ([bool]$item.IsValid) { $validBrush } else { $badBrush }
            $item.Box.BorderBrush = if ([bool]$item.IsValid) { $validBorder } else { $badBorder }
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
        $values = New-Object System.Collections.Generic.List[string]
        foreach ($item in @($Items)) {
            if ($null -eq $item) { continue }
            $itemText = [string]$item
            if ([string]::IsNullOrWhiteSpace($itemText)) { continue }
            [void]$values.Add($itemText)
        }
        $Combo.ItemsSource = $null
        $Combo.Items.Clear()
        foreach ($value in $values) { [void]$Combo.Items.Add($value) }
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
            $Combo.SelectedIndex = $matchIndex
        } catch {
            $Combo.SelectedIndex = -1
        }
    }

    function Test-LocationHierarchyRowMatch {
        param([object]$Row,[string]$Name,[string]$Expected)
        if ([string]::IsNullOrWhiteSpace($Expected)) { return $true }
        return ((Normalize-LocationValue (Get-LocationHierarchyFieldValue $Row $Name)) -eq (Normalize-LocationValue $Expected))
    }

    function Filter-LocationRows {
        param([object[]]$Rows,[string]$City,[string]$Location,[string]$Building,[string]$Floor,[string]$Room)
        $filtered = @($Rows)
        if (-not [string]::IsNullOrWhiteSpace($City)) {
            $nCity = Normalize-LocationValue $City
            $filtered = @($filtered | Where-Object { (Normalize-LocationValue (Get-LocationHierarchyFieldValue $_ 'City')) -eq $nCity })
        }
        if (-not [string]::IsNullOrWhiteSpace($Location)) {
            $nLocation = Normalize-LocationValue $Location
            $filtered = @($filtered | Where-Object { (Normalize-LocationValue (Get-LocationHierarchyFieldValue $_ 'Location')) -eq $nLocation })
        }
        if (-not [string]::IsNullOrWhiteSpace($Building)) {
            $nBuilding = Normalize-LocationValue $Building
            $filtered = @($filtered | Where-Object { (Normalize-LocationValue (Get-LocationHierarchyFieldValue $_ 'Building')) -eq $nBuilding })
        }
        if (-not [string]::IsNullOrWhiteSpace($Floor)) {
            $nFloor = Normalize-LocationValue $Floor
            $filtered = @($filtered | Where-Object { (Normalize-LocationValue (Get-LocationHierarchyFieldValue $_ 'Floor')) -eq $nFloor })
        }
        if (-not [string]::IsNullOrWhiteSpace($Room)) {
            $nRoom = Normalize-LocationValue $Room
            $filtered = @($filtered | Where-Object { (Normalize-LocationValue (Get-LocationHierarchyFieldValue $_ 'Room')) -eq $nRoom })
        }
        return @($filtered)
    }

    function Get-UniqueLocationValues {
        param([object[]]$Rows,[string]$Property,[switch]$Floor)
        $values = @($Rows | ForEach-Object { Get-LocationHierarchyFieldValue $_ $Property } | Where-Object { -not [string]::IsNullOrWhiteSpace([string]$_) })
        if ($Floor) { return @(Sort-LocationFloors -Floors $values) }
        return @($values | Sort-Object -Unique)
    }

    function Populate-LocationCombos {
        param([hashtable]$Ui,[pscustomobject]$Inventory,[string]$ChangedLevel='',[hashtable]$InitialValues=$null)
        if (-not $Ui -or -not $Inventory) { return }
        $script:IsPopulatingLocationCombos = $true
        try {
            $rows = @(Get-LocationHierarchyRows -Inventory $Inventory)

            $cities = @(Get-UniqueLocationValues -Rows $rows -Property 'City')
            if ($ChangedLevel -eq 'City') {
                Set-ControlText -Control $Ui.LocationComboBox -Value ''
                Set-ControlText -Control $Ui.BuildingComboBox -Value ''
                Set-ControlText -Control $Ui.FloorComboBox -Value ''
                Set-ControlText -Control $Ui.RoomComboBox -Value ''
                Set-ControlText -Control $Ui.DepartmentComboBox -Value ''
            } elseif ($ChangedLevel -eq 'Location') {
                Set-ControlText -Control $Ui.BuildingComboBox -Value ''
                Set-ControlText -Control $Ui.FloorComboBox -Value ''
                Set-ControlText -Control $Ui.RoomComboBox -Value ''
                Set-ControlText -Control $Ui.DepartmentComboBox -Value ''
            } elseif ($ChangedLevel -eq 'Building') {
                Set-ControlText -Control $Ui.FloorComboBox -Value ''
                Set-ControlText -Control $Ui.RoomComboBox -Value ''
                Set-ControlText -Control $Ui.DepartmentComboBox -Value ''
            } elseif ($ChangedLevel -eq 'Floor') {
                Set-ControlText -Control $Ui.RoomComboBox -Value ''
                Set-ControlText -Control $Ui.DepartmentComboBox -Value ''
            } elseif ($ChangedLevel -eq 'Room') {
                Set-ControlText -Control $Ui.DepartmentComboBox -Value ''
            }
            $cityText = if ($InitialValues -and $InitialValues.ContainsKey('City')) { [string]$InitialValues.City } else { [string]$Ui.CityComboBox.Text }
            Set-ComboItems -Combo $Ui.CityComboBox -Items $cities -Text $cityText
            $validCity = Get-ValidLocationSelection -Value ([string]$Ui.CityComboBox.Text) -Items $cities

            $locationRows = if ($validCity) { @(Filter-LocationRows -Rows $rows -City $validCity) } else { @($rows) }
            $locations = @(Get-UniqueLocationValues -Rows $locationRows -Property 'Location')
            $locationText = if ($InitialValues -and $InitialValues.ContainsKey('Location')) { [string]$InitialValues.Location } else { [string]$Ui.LocationComboBox.Text }
            if (-not $InitialValues -and -not (Get-ValidLocationSelection -Value $locationText -Items $locations)) { $locationText = '' }
            Set-ComboItems -Combo $Ui.LocationComboBox -Items $locations -Text $locationText
            $validLocation = Get-ValidLocationSelection -Value ([string]$Ui.LocationComboBox.Text) -Items $locations

            $buildingRows = if ($validLocation) { @(Filter-LocationRows -Rows $locationRows -Location $validLocation) } else { @($locationRows) }
            $buildings = @(Get-UniqueLocationValues -Rows $buildingRows -Property 'Building')
            $buildingText = if ($InitialValues -and $InitialValues.ContainsKey('Building')) { [string]$InitialValues.Building } else { [string]$Ui.BuildingComboBox.Text }
            if (-not $InitialValues -and -not (Get-ValidLocationSelection -Value $buildingText -Items $buildings)) { $buildingText = '' }
            Set-ComboItems -Combo $Ui.BuildingComboBox -Items $buildings -Text $buildingText
            $validBuilding = Get-ValidLocationSelection -Value ([string]$Ui.BuildingComboBox.Text) -Items $buildings

            $floorRows = if ($validBuilding) { @(Filter-LocationRows -Rows $buildingRows -Building $validBuilding) } else { @($buildingRows) }
            $floors = @(Get-UniqueLocationValues -Rows $floorRows -Property 'Floor' -Floor)
            $floorText = if ($InitialValues -and $InitialValues.ContainsKey('Floor')) { [string]$InitialValues.Floor } else { [string]$Ui.FloorComboBox.Text }
            if (-not $InitialValues -and -not (Get-ValidLocationSelection -Value $floorText -Items $floors)) { $floorText = '' }
            Set-ComboItems -Combo $Ui.FloorComboBox -Items $floors -Text $floorText
            $validFloor = Get-ValidLocationSelection -Value ([string]$Ui.FloorComboBox.Text) -Items $floors

            $roomRows = if ($validFloor) { @(Filter-LocationRows -Rows $floorRows -Floor $validFloor) } else { @($floorRows) }
            $rooms = @(Get-UniqueLocationValues -Rows $roomRows -Property 'Room')
            $roomText = if ($InitialValues -and $InitialValues.ContainsKey('Room')) { [string]$InitialValues.Room } else { [string]$Ui.RoomComboBox.Text }
            if (-not $InitialValues -and -not (Get-ValidLocationSelection -Value $roomText -Items $rooms)) { $roomText = '' }
            Set-ComboItems -Combo $Ui.RoomComboBox -Items $rooms -Text $roomText
            $validRoom = Get-ValidLocationSelection -Value ([string]$Ui.RoomComboBox.Text) -Items $rooms

            $departmentRows = if ($validRoom) { @(Filter-LocationRows -Rows $roomRows -Room $validRoom) } else { @($roomRows) }
            $departments = @(Get-UniqueLocationValues -Rows $departmentRows -Property 'Department')
            $departmentText = if ($InitialValues -and $InitialValues.ContainsKey('Department')) { [string]$InitialValues.Department } else { [string]$Ui.DepartmentComboBox.Text }
            if (-not $InitialValues -and -not (Get-ValidLocationSelection -Value $departmentText -Items $departments)) { $departmentText = '' }
            Set-ComboItems -Combo $Ui.DepartmentComboBox -Items $departments -Text $departmentText
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
            Model        = 'Dell OptiPlex'
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
            [pscustomobject]@{ Role='Parent'; Type='Tangent'; Name='AO400568'; AssetTag='HSS-8093577'; Serial='C24102M031'; Model='Dell OptiPlex'; SerialForeground='#1F2937'; SerialToolTip=''; RITM='TRP - 26 May 2025'; Retire='31 May 2028' },
            [pscustomobject]@{ Role='Child'; Type='Cart'; Name='AO400568-CRT'; AssetTag='CO09167'; Serial='1896875-0016'; Model='Capsa Cart'; SerialForeground='#1F2937'; SerialToolTip=''; RITM='-'; Retire='-' }
        )

        $nearbyStatusOptions = @('-','Inaccessible - Asset not found','Inaccessible - In storage','Inaccessible - In use by Customer','Inaccessible - Laptop is not onsite','Inaccessible - Other','Inaccessible - Restricted area','Inaccessible - Room locked - Card Swipe','Inaccessible - Room locked - Key Lock','Inaccessible - Under renovation','Inaccessible - User working at home')
        $nearby = @(
            [pscustomobject]@{ HostName='LD065898'; IPAddress='';             Subnet='';       AssetTag='HSS-8077199'; Location='VIHA-DNDR-Duncan Norcr...'; Building='Main Building'; Floor='1'; Room='101 (#6 Charge Cabinet)'; Department='CHS - Community Health Se...'; MaintenanceType='General Rounding'; LastRounded='20 April 2026'; DaysAgo='0';   Status='Inaccessible - Asset not found' },
            [pscustomobject]@{ HostName='LD065911'; IPAddress='10.64.45.232'; Subnet='VPN';    AssetTag='HSS-8077204'; Location='VIHA-DNDR-Duncan Norcr...'; Building='Main Building'; Floor='1'; Room='101 (#7 Charge Cabinet)'; Department='CHS - Community Health Se...'; MaintenanceType='General Rounding'; LastRounded='20 April 2026'; DaysAgo='0';   Status='Inaccessible - Laptop is not onsite' },
            [pscustomobject]@{ HostName='LD062047'; IPAddress='10.64.47.15';  Subnet='VPN';    AssetTag='HSS-1037495'; Location='VIHA-DNDR-Duncan Norcr...'; Building='Main Building'; Floor='1'; Room='101 (#8 Charge Cabinet)'; Department='CHS (Reception)'; MaintenanceType='General Rounding'; LastRounded='20 April 2026'; DaysAgo='0'; Status='Inaccessible - Laptop is not onsite' },
            [pscustomobject]@{ HostName='PC077708'; IPAddress='10.209.233.167';Subnet='Unknown';AssetTag='HSS-1037501'; Location='VIHA-DNDR-Duncan Norcr...'; Building='Main Building'; Floor='1'; Room='102'; Department='CHS - Community Health Se...'; MaintenanceType='General Rounding'; LastRounded='06 March 2026'; DaysAgo='45'; Status='-' },
            [pscustomobject]@{ HostName='LD072236'; IPAddress='10.209.233.47'; Subnet='Unknown';AssetTag='HSS-1037488'; Location='VIHA-DNDR-Duncan Norcr...'; Building='Main Building'; Floor='1'; Room='104 (Chart Room)'; Department='Charting'; MaintenanceType='General Rounding'; LastRounded='20 April 2026'; DaysAgo='0'; Status='-' }
        )
        foreach ($nearbyRow in $nearby) {
            $nearbyRow | Add-Member -NotePropertyName StatusOptions -NotePropertyValue $nearbyStatusOptions -Force
            $nearbyRow | Add-Member -NotePropertyName IsStatusEditable -NotePropertyValue $true -Force
        }

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

    function Set-ConnectivityBadgeUi {
        param([System.Windows.Controls.Border]$Badge,[bool]$IsAvailable)
        if (-not $Badge) { return }
        if ($IsAvailable) { Set-BadgeStyle -Border $Badge -BackgroundHex '#DDF7E5' -ForegroundHex '#15803D' }
        else { Set-BadgeStyle -Border $Badge -BackgroundHex '#FEE2E2' -ForegroundHex '#BE123C' }
    }

    function Set-DeviceNetworkVisibility {
        param([hashtable]$Ui,[bool]$IsVisible)
        $visibility = if ($IsVisible) { 'Visible' } else { 'Collapsed' }
        if ($Ui.DeviceIpText) { $Ui.DeviceIpText.Visibility = $visibility }
        if ($Ui.DeviceSubnetText) { $Ui.DeviceSubnetText.Visibility = $visibility }
    }

    function Set-OnlineStatusUi {
        param(
            [hashtable]$Ui,
            [bool]$IsOnline,
            [Nullable[int]]$LatencyMs,
            [string]$IpAddress='',
            [string]$Subnet='Unknown',
            [string]$CheckedHost=''
        )
        if ($IsOnline) {
            $Ui.DeviceOnlineText.Text = 'Online'
            $Ui.DeviceOnlineText.Foreground = New-Brush '#16A34A'
            $Ui.DeviceOnlineDot.Fill = New-Brush '#16A34A'
            if ($Ui.DeviceStatusIcon) { $Ui.DeviceStatusIcon.Background = New-Brush '#16A34A' }
        }
        else {
            $Ui.DeviceOnlineText.Text = 'Offline'
            $Ui.DeviceOnlineText.Foreground = New-Brush '#BE123C'
            $Ui.DeviceOnlineDot.Fill = New-Brush '#BE123C'
            if ($Ui.DeviceStatusIcon) { $Ui.DeviceStatusIcon.Background = New-Brush '#BE123C' }
        }
        if ($IsOnline -and $LatencyMs -ne $null) { $Ui.DeviceResponseTimeText.Text = "($LatencyMs ms)" }
        elseif (-not [string]::IsNullOrWhiteSpace($CheckedHost)) { $Ui.DeviceResponseTimeText.Text = "($CheckedHost)" }
        else { $Ui.DeviceResponseTimeText.Text = '(No response)' }
        Set-DeviceNetworkVisibility -Ui $Ui -IsVisible:$IsOnline
        if ($IsOnline) {
            if ($Ui.DeviceIpText) { $Ui.DeviceIpText.Text = "IP: $(if ([string]::IsNullOrWhiteSpace($IpAddress)) { 'Unknown' } else { $IpAddress })" }
            if ($Ui.DeviceSubnetText) { $Ui.DeviceSubnetText.Text = "Subnet: $(if ([string]::IsNullOrWhiteSpace($Subnet)) { 'Unknown' } else { $Subnet })" }
        }
        $Ui.LastQueryBadgeText.Text = "Queried $(Get-Date -Format 'HH:mm:ss')"
    }

    function Test-RemoteConnectivity {
        param([string]$HostName,[object]$KnownPingResult=$null,[string]$DataRoot='')
        $result = [pscustomobject]@{ HostName=$HostName; IsOnline=$false; LatencyMs=$null; IpAddress=''; Subnet='Unknown' }
        if ([string]::IsNullOrWhiteSpace($HostName)) { return $result }
        if ($KnownPingResult) {
            $result.IsOnline = [bool]$KnownPingResult.Success
            if ($KnownPingResult.IpAddress -and $KnownPingResult.IpAddress -ne 'Unknown') { $result.IpAddress = [string]$KnownPingResult.IpAddress }
            if ($KnownPingResult.ResponseTime -match '^(\d+(?:\.\d+)?)') { $result.LatencyMs = [int][Math]::Round([double]$Matches[1]) }
            if ($KnownPingResult.Subnet) { $result.Subnet = [string]$KnownPingResult.Subnet }
        }
        if (-not $KnownPingResult) {
            try {
                $reply = Invoke-PingOnce -ComputerName $HostName -TimeoutMilliseconds 1000
                if ($reply -and $reply.Status -eq [System.Net.NetworkInformation.IPStatus]::Success) {
                    $result.IsOnline = $true
                    $latencyValue = Get-FieldValue -Row $reply -Names @('RoundtripTime','ResponseTime','Latency')
                    if (-not [string]::IsNullOrWhiteSpace($latencyValue)) { $result.LatencyMs = [int][Math]::Round([double]$latencyValue) }
                    $result.IpAddress = Get-IPv4AddressFromPingReply -Reply $reply -ComputerName $HostName
                }
            } catch {}
            if (-not [string]::IsNullOrWhiteSpace($result.IpAddress) -and -not [string]::IsNullOrWhiteSpace($DataRoot)) { $result.Subnet = Resolve-SubnetName -IpAddress $result.IpAddress -DataRoot $DataRoot }
        }
        return $result
    }

    function Start-OnlineStatusUpdateAsync {
        param([hashtable]$Ui,[string]$HostName,[string]$QueryToken,[int]$DelayMilliseconds=0)
        if ([string]::IsNullOrWhiteSpace($HostName)) {
            Set-OnlineStatusUi -Ui $Ui -IsOnline:$false -LatencyMs $null
            return
        }
        [System.Threading.Tasks.Task]::Run([Action]{
            try {
                if ($DelayMilliseconds -gt 0) {
                    [System.Threading.Thread]::Sleep($DelayMilliseconds)
                    if (-not [string]::IsNullOrWhiteSpace($QueryToken) -and $script:AppState.CurrentQueryToken -ne $QueryToken) { return }
                }
                $pingResult = Invoke-DevicePing -ComputerName $HostName -DataRoot $script:AppState.DataRoot
                $connectivity = Test-RemoteConnectivity -HostName $HostName -KnownPingResult $pingResult -DataRoot $script:AppState.DataRoot
            }
            catch { $connectivity = [pscustomobject]@{ HostName=$HostName; IsOnline=$false; LatencyMs=$null; IpAddress=''; Subnet='Unknown' } }
            $Ui.MainTabControl.Dispatcher.BeginInvoke([Action]{
                if (-not [string]::IsNullOrWhiteSpace($QueryToken) -and $script:AppState.CurrentQueryToken -ne $QueryToken) { return }
                Set-OnlineStatusUi -Ui $Ui -IsOnline:$connectivity.IsOnline -LatencyMs $connectivity.LatencyMs -IpAddress $connectivity.IpAddress -Subnet $connectivity.Subnet -CheckedHost $connectivity.HostName
                Set-StatusMessage -Ui $Ui -Mode 'PingComplete' -CustomText 'Device found; ping updated'
            }) | Out-Null
        }) | Out-Null
    }

    function Resolve-CurrentPingTarget {
        param([hashtable]$Ui,[object]$Inventory)
        $device = $script:AppState.CurrentDevice
        if (-not $device -and -not [string]::IsNullOrWhiteSpace($Ui.SearchTextBox.Text)) { $device = Find-InventoryMatch -SearchTerm $Ui.SearchTextBox.Text -Inventory $Inventory }
        $parent = if ($device) { Resolve-ParentDevice -Device $device -Inventory $Inventory } else { $null }
        if ($parent -and -not [string]::IsNullOrWhiteSpace($parent.Name)) { return $parent.Name }
        if ($device -and -not [string]::IsNullOrWhiteSpace($device.Name)) { return $device.Name }
        if (-not [string]::IsNullOrWhiteSpace($Ui.SearchTextBox.Text)) { return $Ui.SearchTextBox.Text.Trim() }
        return $null
    }

    function Invoke-CurrentDevicePing {
        param([hashtable]$Ui,[object]$Inventory,[string]$DataRoot,[switch]$StartContinuous)
        $target = Resolve-CurrentPingTarget -Ui $Ui -Inventory $Inventory
        if ([string]::IsNullOrWhiteSpace($target)) { throw 'Enter or query a device before using Ping.' }

        $pingResult = Invoke-DevicePing -ComputerName $target -DataRoot $DataRoot
        $connectivity = Test-RemoteConnectivity -HostName $target -KnownPingResult $pingResult -DataRoot $DataRoot
        Set-OnlineStatusUi -Ui $Ui -IsOnline:$connectivity.IsOnline -LatencyMs $connectivity.LatencyMs -IpAddress $connectivity.IpAddress -Subnet $connectivity.Subnet -CheckedHost $connectivity.HostName

        if ($StartContinuous) {
            Start-ContinuousPingWindow -Target $(if ($pingResult.IpAddress -and $pingResult.IpAddress -ne 'Unknown') { $pingResult.IpAddress } else { $target })
            Set-StatusMessage -Ui $Ui -Mode 'PingComplete' -CustomText 'Continuous ping started; device status updated'
        }
        else {
            Set-StatusMessage -Ui $Ui -Mode 'PingComplete' -CustomText 'Ping complete; device status updated'
        }
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
        param([hashtable]$Ui,[bool]$ClearNearby=$true)
        foreach ($name in @('DetectedType','HostName','AssetTag','Serial','Parent','Ritm','Retire')) { Set-DisplayText -Ui $Ui -BaseName $name -Value '' }
        Set-ControlText -Control $Ui.SelectedDeviceText -Value ''
        $Ui.LastRoundedLabelText.Foreground = New-Brush '#64748B'
        Set-ControlText -Control $Ui.LastRoundedText -Value ''
        $Ui.LastRoundedText.Foreground = New-Brush '#64748B'
        $Ui.LastRoundedContainer.Background = New-Brush '#F8FAFC'
        $Ui.LastRoundedContainer.BorderBrush = New-Brush '#D9E1EA'
        $Ui.LastRoundedAttentionBadge.Visibility = 'Collapsed'
        foreach ($box in @($Ui.CityTextBox,$Ui.LocationTextBox,$Ui.BuildingTextBox,$Ui.FloorTextBox,$Ui.RoomTextBox,$Ui.DepartmentTextBox)) { Set-ControlText -Control $box -Value ''; $box.Background = New-Brush '#CCF2D3'; $box.BorderBrush = New-Brush '#7CE0A6' }
        $Ui.AssociatedDevicesDataGrid.ItemsSource = @()
        if ($ClearNearby) {
            Set-ControlText -Control $Ui.NearbyScopeSummaryText -Value 'Nearby disabled'
            $Ui.NearbyDataGrid.ItemsSource = @()
        } else {
            Update-NearbySummary -Ui $Ui
        }
        Set-ControlText -Control $Ui.DeviceOnlineText -Value 'Ready'
        $Ui.DeviceOnlineText.Foreground = New-Brush '#64748B'
        $Ui.DeviceOnlineDot.Fill = New-Brush '#94A3B8'
        Set-ControlText -Control $Ui.DeviceResponseTimeText -Value ''
        Set-ControlText -Control $Ui.LastQueryBadgeText -Value 'Awaiting query'
        Set-ControlText -Control $Ui.DeviceIpText -Value 'IP: Unknown'
        Set-ControlText -Control $Ui.DeviceSubnetText -Value 'Subnet: Unknown'
        Set-DeviceNetworkVisibility -Ui $Ui -IsVisible:$false
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
        param([hashtable]$Ui,[pscustomobject]$CurrentDevice,[pscustomobject]$Inventory,[hashtable]$RoundingByAssetTag,[string]$ResolvedXamlPath)
        $parentDevice = Resolve-ParentDevice -Device $CurrentDevice -Inventory $Inventory
        if (-not $parentDevice -or $parentDevice.DetectedType -ne 'Computer') {
            [System.Windows.MessageBox]::Show('No parent computer was found in the Computers CSV data. RoundingEvents.csv only accepts parent computer records.', 'Save Event') | Out-Null
            return $false
        }

        Ensure-OutputFolder -ResolvedXamlPath $ResolvedXamlPath
        $csvPath = Get-RoundingEventsPath -ResolvedXamlPath $ResolvedXamlPath
        $row = [pscustomobject]([ordered]@{
            Timestamp=(Get-Date).ToString('yyyy-MM-dd HH:mm:ss')
            AssetTag=$parentDevice.AssetTag; Name=$parentDevice.Name; Serial=$parentDevice.Serial
            City=$parentDevice.City; Location=$parentDevice.Location; Building=$parentDevice.Building
            Floor=$parentDevice.Floor; Room=$parentDevice.Room; CheckStatus=$Ui.CheckStatusComboBox.Text
            RoundingMinutes=(Get-RoundingMinutes -Ui $Ui); CableMgmtOK=$(if($Ui.ValidateCableCheckBox.IsChecked){'Yes'}else{'No'})
            CablingNeeded=$(if($Ui.CablingNeededCheckBox.IsChecked){'Yes'}else{'No'})
            LabelOK=$(if($Ui.LabelMonitorCheckBox.IsChecked){'Yes'}else{'No'}); CartOK=$(if($Ui.PhysicalCartCheckBox.IsChecked){'Yes'}else{'No'})
            PeripheralsOK=$(if($Ui.ValidatePeripheralsCheckBox.IsChecked){'Yes'}else{'No'})
            MaintenanceType=$Ui.MaintenanceTypeComboBox.Text; Department=$parentDevice.Department
            RoundingUrl=(Get-RoundingUrlForDevice -CurrentDevice $parentDevice -RoundingByAssetTag $RoundingByAssetTag); Comments=$Ui.CommentsTextBox.Text
            Rounded=$(if($script:ManualRoundUsed){'Yes'}else{'No'})
        })
        Add-RoundingCsvRow -Path $csvPath -Row $row
        if ($script:RoundingPlan) { Update-RoundingPlanBadges -Ui $Ui -ResolvedXamlPath $ResolvedXamlPath -Plan $script:RoundingPlan }
        $script:ManualRoundUsed = $false
        [System.Windows.MessageBox]::Show("Saved parent computer event to:`n$csvPath", 'Save Event') | Out-Null
        return $true
    }

    function Save-NearbyEvents {
        param([hashtable]$Ui,[pscustomobject]$Inventory,[hashtable]$RoundingByAssetTag,[string]$ResolvedXamlPath)
        if (-not $Ui -or -not $Ui.NearbyDataGrid) { return }
        Ensure-OutputFolder -ResolvedXamlPath $ResolvedXamlPath
        $csvPath = Get-RoundingEventsPath -ResolvedXamlPath $ResolvedXamlPath
        Load-NearbyRoundingEvents -ResolvedXamlPath $ResolvedXamlPath

        $nearbyRows = @($Ui.NearbyDataGrid.ItemsSource)
        if ($nearbyRows.Count -eq 0) {
            [System.Windows.MessageBox]::Show('There are no Nearby rows to save.', 'Nearby Save') | Out-Null
            return
        }

        $saved = 0
        foreach ($item in $nearbyRows) {
            if (-not $item) { continue }
            if ($item.PSObject.Properties.Name -contains 'IsStatusEditable' -and -not [bool]$item.IsStatusEditable) { continue }
            $status = [string]$item.Status
            if ([string]::IsNullOrWhiteSpace($status) -or $status -eq '-') { continue }

            $assetKey = if ($item.AssetTag) { $item.AssetTag.Trim().ToUpperInvariant() } else { '' }
            if (-not [string]::IsNullOrWhiteSpace($assetKey) -and $script:AppState.NearbyRoundedTodayAssetTags -and $script:AppState.NearbyRoundedTodayAssetTags.Contains($assetKey)) { continue }

            $device = $null
            if ($item.PSObject.Properties.Name -contains 'Device') { $device = $item.Device }
            if (-not $device -and $Inventory -and -not [string]::IsNullOrWhiteSpace($assetKey) -and $Inventory.IndexByAsset -and $Inventory.IndexByAsset.ContainsKey($assetKey)) {
                $device = $Inventory.IndexByAsset[$assetKey]
            }
            if (-not $device) {
                $device = [pscustomobject]@{
                    Name=$item.HostName; AssetTag=$item.AssetTag; Serial=''; City=''
                    Location=$item.Location; Building=$item.Building; Floor=$item.Floor; Room=$item.Room
                    Department=$item.Department; DetectedType='Computer'
                }
            }

            $row = [pscustomobject]([ordered]@{
                Timestamp=(Get-Date).ToString('yyyy-MM-dd HH:mm:ss')
                AssetTag=$item.AssetTag; Name=$item.HostName; Serial=''; City='Duncan'; Location=$item.Location; Building=$item.Building; Floor=$item.Floor; Room=$item.Room
                CheckStatus=$(if ([string]::IsNullOrWhiteSpace($item.Status) -or $item.Status -eq '-') { 'Complete' } else { $item.Status })
                RoundingMinutes=3; CableMgmtOK='No'; CablingNeeded='No'; LabelOK='No'; CartOK='No'; PeripheralsOK='No'
                MaintenanceType=$item.MaintenanceType; Department=$item.Department
                RoundingUrl=(Get-RoundingUrlForDevice -CurrentDevice $device -RoundingByAssetTag $RoundingByAssetTag)
                Comments=''; Rounded='No'
            })
            Add-RoundingCsvRow -Path $csvPath -Row $row
            $saved++
        }

        if ($saved -gt 0) {
            if ($script:RoundingPlan) { Update-RoundingPlanBadges -Ui $Ui -ResolvedXamlPath $ResolvedXamlPath -Plan $script:RoundingPlan }
            Load-NearbyRoundingEvents -ResolvedXamlPath $ResolvedXamlPath
            Update-NearbyRows -Ui $Ui -Inventory $Inventory -ResolvedXamlPath $ResolvedXamlPath
            [System.Windows.MessageBox]::Show("Saved $saved nearby rounding event(s) to:`n$csvPath", 'Nearby Save') | Out-Null
        } else {
            [System.Windows.MessageBox]::Show('Nothing to save. Pick a Nearby status first.', 'Nearby Save') | Out-Null
        }
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
            $editablePairs = @(
                [pscustomobject]@{ Combo=$Ui.CityComboBox; TextBox=$Ui.CityTextBox },
                [pscustomobject]@{ Combo=$Ui.LocationComboBox; TextBox=$Ui.LocationTextBox },
                [pscustomobject]@{ Combo=$Ui.BuildingComboBox; TextBox=$Ui.BuildingTextBox },
                [pscustomobject]@{ Combo=$Ui.FloorComboBox; TextBox=$Ui.FloorTextBox },
                [pscustomobject]@{ Combo=$Ui.RoomComboBox; TextBox=$Ui.RoomTextBox },
                [pscustomobject]@{ Combo=$Ui.DepartmentComboBox; TextBox=$Ui.DepartmentTextBox }
            )
            foreach ($pair in $editablePairs) {
                if (-not $pair.Combo -or -not $pair.TextBox) { continue }
                $pair.Combo.IsEditable = $true
                $pair.Combo.Text = $pair.TextBox.Text
            }
            $initialValues = @{
                City = [string]$Ui.CityTextBox.Text
                Location = [string]$Ui.LocationTextBox.Text
                Building = [string]$Ui.BuildingTextBox.Text
                Floor = [string]$Ui.FloorTextBox.Text
                Room = [string]$Ui.RoomTextBox.Text
                Department = [string]$Ui.DepartmentTextBox.Text
            }
            Populate-LocationCombos -Ui $Ui -Inventory $Inventory -InitialValues $initialValues
        }
    }

    function Get-LocationUserAddsPath {
        param([pscustomobject]$Inventory)
        $folder = if ($Inventory -and -not [string]::IsNullOrWhiteSpace($Inventory.SiteFolderPath) -and (Test-Path -LiteralPath $Inventory.SiteFolderPath)) { $Inventory.SiteFolderPath } elseif ($Inventory) { $Inventory.DataRoot } else { $null }
        if ([string]::IsNullOrWhiteSpace($folder)) { return $null }
        $existing = @(Get-ChildItem -LiteralPath $folder -File -Filter 'LocationMaster-UserAdds*.csv' -ErrorAction SilentlyContinue | Sort-Object Name | Select-Object -First 1)
        if ($existing.Count -gt 0) { return $existing[0].FullName }
        $suffix = if ($Inventory -and -not [string]::IsNullOrWhiteSpace($Inventory.SiteFolderPath)) { ' - ' + (Split-Path -Leaf $Inventory.SiteFolderPath) } else { '' }
        return (Join-Path $folder ("LocationMaster-UserAdds$suffix.csv"))
    }

    function Add-LocationUserAddRow {
        param([pscustomobject]$Inventory,[string]$City,[string]$Location,[string]$Building,[string]$Floor,[string]$Room,[string]$Department)
        if (-not $Inventory) { return }
        if ([string]::IsNullOrWhiteSpace($City) -and [string]::IsNullOrWhiteSpace($Location) -and [string]::IsNullOrWhiteSpace($Building) -and [string]::IsNullOrWhiteSpace($Floor) -and [string]::IsNullOrWhiteSpace($Room) -and [string]::IsNullOrWhiteSpace($Department)) { return }
        $newRow = New-LocationHierarchyRow -City $City -Location $Location -Building $Building -Floor $Floor -Room $Room -Department $Department
        $newKey = '{0}|{1}|{2}|{3}|{4}|{5}' -f (Normalize-LocationValue $City),(Normalize-LocationValue $Location),(Normalize-LocationValue $Building),(Normalize-LocationValue $Floor),(Normalize-LocationValue $Room),(Normalize-LocationValue $Department)
        foreach ($row in @(Get-LocationHierarchyRows -Inventory $Inventory)) {
            $key = '{0}|{1}|{2}|{3}|{4}|{5}' -f (Normalize-LocationValue $row.City),(Normalize-LocationValue $row.Location),(Normalize-LocationValue $row.Building),(Normalize-LocationValue $row.Floor),(Normalize-LocationValue $row.Room),(Normalize-LocationValue $row.Department)
            if ($key -eq $newKey) { return }
        }
        $path = Get-LocationUserAddsPath -Inventory $Inventory
        if (-not [string]::IsNullOrWhiteSpace($path)) {
            $dir = Split-Path -Path $path -Parent
            if ($dir -and -not (Test-Path -LiteralPath $dir)) { New-Item -Path $dir -ItemType Directory -Force | Out-Null }
            @($newRow) | Export-Csv -Path $path -NoTypeInformation -Append:((Test-Path -LiteralPath $path)) -Force
        }
        $Inventory.Locations = @($Inventory.Locations) + $newRow
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
            if ($script:AppState.Inventory) {
                Add-LocationUserAddRow -Inventory $script:AppState.Inventory -City $Ui.CityTextBox.Text -Location $Ui.LocationTextBox.Text -Building $Ui.BuildingTextBox.Text -Floor $Ui.FloorTextBox.Text -Room $Ui.RoomTextBox.Text -Department $Ui.DepartmentTextBox.Text
                Set-LocationValidationStyle -Ui $Ui -Inventory $script:AppState.Inventory
            }
        }
    }

    $resolvedXamlPath = (Resolve-Path -LiteralPath $XamlPath).Path
    $window = ConvertFrom-XamlFile -Path $resolvedXamlPath
    Set-WindowIconFromFile -Window $window -ResolvedXamlPath $resolvedXamlPath

    $ui = Get-NamedControls -Window $window -Names @(
        'SearchTextBox','QueryButton','PingButton','LiveDetailsButton','MonitorLabelButton',
        'MainTabControl','SystemTab','NearbyTab','SelectedDeviceText','DeviceStatusIcon','DeviceOnlineText','DeviceOnlineDot','DeviceResponseTimeText','LastQueryBadgeText',
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
        'NearbyScopeSummaryText','RebuildNearbyButton','PingAllButton','IsolateNearbyButton','ClearNearbyButton',
        'NearbyDataGrid','NearbySaveButton','ShowAllNearbyButton',
        'ShowAllNearbyCheckBox','TodaysRoundedCheckBox','ExcludedCheckBox','RecentlyRoundedCheckBox','CriticalClinicalCheckBox',
        'DataPathText','OutputPathText','DaysPerWeekBadge','DaysPerWeekBadgeText','TodayBadge','TodayBadgeText','ThisWeekBadge','ThisWeekBadgeText','RemainingPerDayBadge','RemainingPerDayBadgeText','StatusMessageBadge','DataFileBadge','DataFileBadgeText','DeviceIpText','DeviceSubnetText'
    )
    $ui.Window = $window

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
    $dataFiles = Get-DataFileInfo -ResolvedXamlPath $resolvedXamlPath -SiteFolderPath $siteFolderPath
    $script:AppState = [pscustomobject]@{ LastStatusMode='Found'; SampleData=$null; CurrentDevice=$null; CurrentQueryToken=''; Inventory=$inventory; SelectedSiteName=$siteName; SelectedSummaryDevice=$null; SelectedSummaryParent=$null; DataRoot=$dataRoot; DataFiles=$dataFiles; ActiveNearbyScopes=(New-Object 'System.Collections.Generic.HashSet[string]') }

    Clear-WindowData -Ui $ui
    Set-RoundingMinutes -Ui $ui -Minutes 3
    Increment-Fonts -Root $window
    Set-StatusMessage -Ui $ui -Mode 'Warning' -CustomText 'Ready. Enter a device and click Query.'
    Update-DataFileBadge -Ui $ui -DataFiles $dataFiles
    Toggle-LocationEditMode -Ui $ui -IsEditing:$false
    Register-SummaryClipboardCopy -Ui $ui
    if ($ui.NearbyDataGrid) {
        $nearbyWheelHandler = [System.Windows.Input.MouseWheelEventHandler]{
            param($sender, $e)
            $scrollViewer = Find-VisualChildByType -Root $ui.NearbyDataGrid -Type ([System.Windows.Controls.ScrollViewer])
            if ($null -eq $scrollViewer) { return }
            if ($e.Delta -lt 0) { $scrollViewer.LineDown() } else { $scrollViewer.LineUp() }
            $e.Handled = $true
        }
        $ui.NearbyDataGrid.AddHandler([System.Windows.UIElement]::PreviewMouseWheelEvent, $nearbyWheelHandler, $true)
        $ui.NearbyDataGrid.AddHandler([System.Windows.UIElement]::MouseWheelEvent, $nearbyWheelHandler, $true)
        Initialize-NearbyContextMenu -Ui $ui -DataRoot $dataRoot
    }

    $window.Title = "New Inventory Tool - $siteName"
    Set-ControlText -Control $ui.DataPathText -Value "Data: $dataRoot (Site: $siteName)"
    Set-ControlText -Control $ui.OutputPathText -Value "Output: $(Get-OutputFolder -ResolvedXamlPath $resolvedXamlPath)"
    $ui.DaysPerWeekBadge.Add_MouseLeftButtonUp({ Ensure-RoundingPlan -Ui $ui -Window $window -ResolvedXamlPath $resolvedXamlPath -Force:$true })
    $ui.DataFileBadge.Add_MouseLeftButtonUp({ Show-DataFileInfo -Ui $ui })

    foreach ($combo in @($ui.CityComboBox,$ui.LocationComboBox,$ui.BuildingComboBox,$ui.FloorComboBox,$ui.RoomComboBox,$ui.DepartmentComboBox)) {
        $combo.Items.Clear()
        Set-ControlText -Control $combo -Value ''
    }

    $script:IsUpdatingNearbyFilters = $false
    $refreshNearbyFromFilters = {
        if ($script:IsUpdatingNearbyFilters) { return }
        Update-NearbyRows -Ui $ui -Inventory $script:AppState.Inventory -ResolvedXamlPath $resolvedXamlPath
    }
    foreach ($checkBox in @($ui.TodaysRoundedCheckBox,$ui.ExcludedCheckBox,$ui.RecentlyRoundedCheckBox,$ui.CriticalClinicalCheckBox)) {
        $checkBox.Add_Click($refreshNearbyFromFilters)
    }
    if ($ui.ShowAllNearbyButton) { $ui.ShowAllNearbyButton.Add_Click({ if ($ui.ShowAllNearbyCheckBox) { $ui.ShowAllNearbyCheckBox.IsChecked = $true; $ui.ShowAllNearbyCheckBox.RaiseEvent((New-Object System.Windows.RoutedEventArgs([System.Windows.Controls.Primitives.ButtonBase]::ClickEvent))) } }) }

    $ui.ShowAllNearbyCheckBox.Add_Click({
        try {
            $script:IsUpdatingNearbyFilters = $true
            if ($ui.ShowAllNearbyCheckBox.IsChecked) {
                $ui.TodaysRoundedCheckBox.IsChecked = $true
                $ui.ExcludedCheckBox.IsChecked = $true
                $ui.RecentlyRoundedCheckBox.IsChecked = $true
                $ui.CriticalClinicalCheckBox.IsChecked = $true
            } else {
                $ui.TodaysRoundedCheckBox.IsChecked = $false
                $ui.ExcludedCheckBox.IsChecked = $false
                $ui.RecentlyRoundedCheckBox.IsChecked = $true
                $ui.CriticalClinicalCheckBox.IsChecked = $false
            }
        } finally {
            $script:IsUpdatingNearbyFilters = $false
        }
        Update-NearbyRows -Ui $ui -Inventory $script:AppState.Inventory -ResolvedXamlPath $resolvedXamlPath
    })

    $ui.CityComboBox.Add_SelectionChanged({ if (-not $script:IsPopulatingLocationCombos) { Populate-LocationCombos -Ui $ui -Inventory $script:AppState.Inventory -ChangedLevel 'City' } })
    $ui.LocationComboBox.Add_SelectionChanged({ if (-not $script:IsPopulatingLocationCombos) { Populate-LocationCombos -Ui $ui -Inventory $script:AppState.Inventory -ChangedLevel 'Location' } })
    $ui.BuildingComboBox.Add_SelectionChanged({ if (-not $script:IsPopulatingLocationCombos) { Populate-LocationCombos -Ui $ui -Inventory $script:AppState.Inventory -ChangedLevel 'Building' } })
    $ui.FloorComboBox.Add_SelectionChanged({ if (-not $script:IsPopulatingLocationCombos) { Populate-LocationCombos -Ui $ui -Inventory $script:AppState.Inventory -ChangedLevel 'Floor' } })
    $ui.RoomComboBox.Add_SelectionChanged({ if (-not $script:IsPopulatingLocationCombos) { Populate-LocationCombos -Ui $ui -Inventory $script:AppState.Inventory -ChangedLevel 'Room' } })

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
        $existingNearby = @()
        try { $existingNearby = @($ui.NearbyDataGrid.ItemsSource) } catch {}
        $script:AppState.SampleData = [pscustomobject]@{ Device=$match; Associated=$associated; Nearby=$existingNearby }
        $script:AppState.CurrentQueryToken = [guid]::NewGuid().ToString('N')
        Set-PrimaryDeviceBindings -Ui $ui -Device $match -Inventory $inventory
        $ui.AssociatedDevicesDataGrid.ItemsSource = $associated
        Set-ControlText -Control $ui.DeviceOnlineText -Value 'Checking...'
        $ui.DeviceOnlineText.Foreground = New-Brush '#64748B'
        $ui.DeviceOnlineDot.Fill = New-Brush '#94A3B8'
        if ($ui.DeviceStatusIcon) { $ui.DeviceStatusIcon.Background = New-Brush '#94A3B8' }
        Set-ControlText -Control $ui.DeviceResponseTimeText -Value ''
        Set-ControlText -Control $ui.DeviceIpText -Value 'IP: Checking...'
        Set-ControlText -Control $ui.DeviceSubnetText -Value 'Subnet: Checking...'
        Set-DeviceNetworkVisibility -Ui $ui -IsVisible:$false
        Set-ControlText -Control $ui.LastQueryBadgeText -Value "Queried $(Get-Date -Format 'HH:mm:ss')"
        Set-StatusMessage -Ui $ui -Mode 'Found'
        Start-OnlineStatusUpdateAsync -Ui $ui -HostName (Resolve-CurrentPingTarget -Ui $ui -Inventory $script:AppState.Inventory) -QueryToken $script:AppState.CurrentQueryToken
    })

    function Reset-RoundingFormForNextScan {
        param([hashtable]$Ui)
        foreach ($cb in @($Ui.ValidateCableCheckBox,$Ui.LabelMonitorCheckBox,$Ui.ValidatePeripheralsCheckBox,$Ui.CablingNeededCheckBox,$Ui.PhysicalCartCheckBox,$Ui.AddDeviceToTrackerCheckBox)) { $cb.IsChecked = $false }
        Set-ControlText -Control $Ui.CheckStatusComboBox -Value '-'
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
            Clear-WindowData -Ui $ui -ClearNearby:$false
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
        try {
            Invoke-CurrentDevicePing -Ui $ui -Inventory $script:AppState.Inventory -DataRoot $dataRoot -StartContinuous
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
    $ui.EditLocationButton.Add_Click({
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
            [System.Windows.MessageBox]::Show(("{0}`n`n{1}" -f $_.Exception.Message, $_.ScriptStackTrace), 'Edit Location') | Out-Null
        }
    })
    $ui.CancelEditLocationButton.Add_Click({
        param($sender,$e)
        try {
            Toggle-LocationEditMode -Ui $ui -IsEditing:$false
            Set-LocationValidationStyle -Ui $ui -Inventory $script:AppState.Inventory
        } catch {
            [System.Windows.MessageBox]::Show(("{0}`n`n{1}" -f $_.Exception.Message, $_.ScriptStackTrace), 'Edit Location') | Out-Null
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
        Set-ControlText -Control $ui.CheckStatusComboBox -Value '-'
    })
    $ui.SaveEventButton.Add_Click({
        if (-not $script:AppState.CurrentDevice) { return }
        $roundingTimer.Stop()
        $script:RoundingStartTimeUtc = $null
        if ($ui.MaintenanceTypeComboBox.Text -eq 'Excluded') {
            [System.Windows.MessageBox]::Show("This device is marked as Excluded. Enable 'Excluded' to log rounding.","Save Event") | Out-Null
            return
        }
        $saved = Save-RoundingEvent -Ui $ui -CurrentDevice $script:AppState.CurrentDevice -Inventory $script:AppState.Inventory -RoundingByAssetTag $roundingByAssetTag -ResolvedXamlPath $resolvedXamlPath
        if (-not $saved) { return }
        $nearbyScopeDevice = Resolve-ParentDevice -Device $script:AppState.CurrentDevice -Inventory $script:AppState.Inventory
        if (-not $nearbyScopeDevice) { $nearbyScopeDevice = $script:AppState.CurrentDevice }
        [void](Add-NearbyScope -Device $nearbyScopeDevice)
        Update-NearbyRows -Ui $ui -Inventory $script:AppState.Inventory
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
    $ui.RebuildNearbyButton.Add_Click({ Update-NearbyRows -Ui $ui -Inventory $script:AppState.Inventory })
    $ui.ClearNearbyButton.Add_Click({ Ensure-NearbyState; $script:AppState.ActiveNearbyScopes.Clear(); $ui.NearbyDataGrid.ItemsSource = @(); Update-NearbySummary -Ui $ui })
    $ui.PingAllButton.Add_Click({ Set-ControlText -Control $ui.NearbyScopeSummaryText -Value 'Nearby disabled' })
    $ui.NearbySaveButton.Add_Click({ Save-NearbyEvents -Ui $ui -Inventory $script:AppState.Inventory -RoundingByAssetTag $roundingByAssetTag -ResolvedXamlPath $resolvedXamlPath })
    $ui.AssociatedDevicesDataGrid.Add_MouseDoubleClick({
        $row = $ui.AssociatedDevicesDataGrid.SelectedItem
        if (-not $row) { return }
        $device = $null
        if ($row.PSObject.Properties.Name -contains 'Device') { $device = $row.Device }
        if (-not $device) {
            $device = [pscustomobject]@{
                DetectedType=$row.Type; Name=$row.Name; AssetTag=$row.AssetTag; Serial=$row.Serial; Model=$row.Model; Parent='(n/a)'; RITM=$row.RITM; RetireDate=$row.Retire
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
    $window.Add_Loaded({ Ensure-RoundingPlan -Ui $ui -Window $window -ResolvedXamlPath $resolvedXamlPath -Force:$false })

    [void]$window.ShowDialog()
}
catch {
    Show-StartupError -Exception $_.Exception -ScriptPath $PSCommandPath
    throw
}
