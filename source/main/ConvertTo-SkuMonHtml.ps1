function ConvertTo-SkuMonHtml {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory, ValueFromPipeline)]
        [psobject[]]$InputObject,

        [Parameter()]
        [string]$ReportTitle = 'Microsoft 365 License Availability Report',

        [Parameter()]
        [string]$OrganizationName
    )

    begin {
        $items = [System.Collections.Generic.List[object]]::new()
        $LogoSize = 24
        $LegendLogoSize = 16
        $BarWidth = 150
        $BarHeight = 10

        $barColors = @{
            'Normal'  = '#6B8F71' # Warmer muted green
            'Warning' = '#C0392B' # Red for warning
            'Ignore'  = '#7F8C8D' # Gray for ignored SKUs
        }

        function Get-HtmlEncodedText {
            param (
                [AllowNull()]
                [object]$Value
            )

            if ($null -eq $Value) {
                return ''
            }

            return [System.Net.WebUtility]::HtmlEncode([string]$Value)
        }

        function Format-HtmlNumber {
            param (
                [AllowNull()]
                [object]$Value
            )

            if ($null -eq $Value) {
                return '0'
            }

            return '{0:N0}' -f [double]$Value
        }

        function Get-ExceptionNumberStyle {
            param (
                [Parameter(Mandatory)]
                [string]$State,

                [AllowNull()]
                [object]$Value
            )

            $number = 0

            if ($null -ne $Value) {
                $number = [int]$Value
            }

            if ($number -le 0) {
                return ''
            }

            switch ($State) {
                'LockedOut' {
                    return 'background-color:#FDE7E9;color:#A4262C;font-weight:bold;'
                }
                'Suspended' {
                    return 'background-color:#FCE4D6;color:#A64100;font-weight:bold;'
                }
                'Warning' {
                    return 'background-color:#FFF4CE;color:#8A5A00;font-weight:bold;'
                }
                default {
                    return ''
                }
            }
        }
    }

    process {
        foreach ($item in $InputObject) {
            $items.Add($item)
        }
    }

    end {
        if (-not $items.Count) {
            throw "No SkuMonData objects were provided."
        }

        # Resolve resource folder.
        $module = Get-Module PsSkuMon365

        if (-not $module) {
            throw "Module 'PsSkuMon365' is not loaded. Unable to resolve the resource folder."
        }

        $ResourceFolder = Join-Path (Split-Path $module.Path -Parent) 'resource'

        # Load CSS.
        $cssPath = Join-Path $ResourceFolder 'style.css'
        $css = Get-Content $cssPath -Raw

        # Load PNG images as base64.
        $normalButtonPath = Join-Path $ResourceFolder 'normal.png'
        $ignoreButtonPath = Join-Path $ResourceFolder 'ignore.png'
        $warningButtonPath = Join-Path $ResourceFolder 'warning.png'

        $normalButtonBase64 = [convert]::ToBase64String([System.IO.File]::ReadAllBytes($normalButtonPath))
        $ignoreButtonBase64 = [convert]::ToBase64String([System.IO.File]::ReadAllBytes($ignoreButtonPath))
        $warningButtonBase64 = [convert]::ToBase64String([System.IO.File]::ReadAllBytes($warningButtonPath))

        # Organization info.
        if (-not $OrganizationName) {
            $org = Get-MgOrganization -ErrorAction Stop
            $OrganizationName = Get-HtmlEncodedText $org.DisplayName
        }
        else {
            $OrganizationName = Get-HtmlEncodedText $OrganizationName
        }

        # Time info.
        $tzInfo = [System.TimeZoneInfo]::Local
        $tz = ($tzInfo.DisplayName -split ' ')[0]
        $today = Get-Date -Format f

        $html = @()

        # Header.
        $html += '<html><head><title>' + (Get-HtmlEncodedText $ReportTitle) + '</title>'
        $html += '<style type="text/css">'
        $html += $css
        $html += '</style></head><body>'

        # Title section.
        $html += '<table id="tbl">'
        $html += '<tr><td class="head"></td></tr>'
        $html += '<tr><th class="section">Microsoft 365 Licenses Report</th></tr>'
        $html += '<tr><td class="head"><b>' + $OrganizationName + '</b><br>' + $today + ' ' + $tz + '</td></tr>'
        $html += '</table>'

        # Subscription Health exception section.
        # This section intentionally shows only SKUs that need attention.
        $healthItems = @(
            $items | Where-Object {
                ([int]$_.Warning -gt 0) -or
                ([int]$_.Suspended -gt 0) -or
                ([int]$_.LockedOut -gt 0)
            } | Sort-Object `
            @{ Expression = { [int]$_.LockedOut }; Descending = $true },
            @{ Expression = { [int]$_.Suspended }; Descending = $true },
            @{ Expression = { [int]$_.Warning }; Descending = $true },
            SkuName
        )

        if ($healthItems.Count -gt 0) {
            $html += '<table id="tbl">'
            $html += '<tr><td class="head" colspan="8"></td></tr>'
            $html += '<tr><th class="section" colspan="8" style="border-top: 2px solid #CCC;">Subscription Health</th></tr>'
            $html += '<tr>'
            $html += '<td colspan="6" class="head">Only SKUs with Warning, Suspended, or Locked Out units are shown.</td>'
            $html += '</tr>'
            $html += '<tr style="border-top: 2px solid #CCC;">'
            $html += '<td>Name</td>'
            $html += '<td>Total Usable</td>'
            $html += '<td>Used</td>'
            $html += '<td>Free</td>'
            $html += '<td>Enabled</td>'
            $html += '<td>Warning</td>'
            $html += '<td>Suspended</td>'
            $html += '<td>Locked Out</td>'
            $html += '</tr>'

            foreach ($item in $healthItems) {
                $skuName = Get-HtmlEncodedText $item.SkuName

                $total = Format-HtmlNumber $item.Total
                $assigned = Format-HtmlNumber $item.Assigned
                $enabled = Format-HtmlNumber $item.Enabled
                $free = Format-HtmlNumber $item.Available
                $warning = Format-HtmlNumber $item.Warning
                $suspended = Format-HtmlNumber $item.Suspended
                $lockedOut = Format-HtmlNumber $item.LockedOut

                $warningStyle = Get-ExceptionNumberStyle -State 'Warning' -Value $item.Warning
                $suspendedStyle = Get-ExceptionNumberStyle -State 'Suspended' -Value $item.Suspended
                $lockedOutStyle = Get-ExceptionNumberStyle -State 'LockedOut' -Value $item.LockedOut

                $html += '<tr>'
                $html += '<td valign="middle" style="vertical-align:middle;font-weight:bold;">' + $skuName + '</td>'
                $html += '<td valign="middle" style="vertical-align:middle;text-align:right;">' + $total + '</td>'
                $html += '<td valign="middle" style="vertical-align:middle;text-align:right;">' + $assigned + '</td>'
                $html += '<td valign="middle" style="vertical-align:middle;text-align:right;">' + $free + '</td>'
                $html += '<td valign="middle" style="vertical-align:middle;text-align:right;">' + $enabled + '</td>'
                $html += '<td valign="middle" style="vertical-align:middle;text-align:right;' + $warningStyle + '">' + $warning + '</td>'
                $html += '<td valign="middle" style="vertical-align:middle;text-align:right;' + $suspendedStyle + '">' + $suspended + '</td>'
                $html += '<td valign="middle" style="vertical-align:middle;text-align:right;' + $lockedOutStyle + '">' + $lockedOut + '</td>'
                $html += '</tr>'
            }

            $html += '<tr><td class="head" colspan="6"></td></tr>'
            $html += '</table>'
        }

        $html += '<table id="tbl" cellpadding="0" cellspacing="0" border="0">'
        $html += '<tr>'

        $html += '<th class="section" colspan="4" align="left" width="60%" valign="middle" style="vertical-align:middle;">'
        $html += 'License Utilization'
        $html += '</th>'

        $html += '<td align="right" width="40%" valign="middle" style="vertical-align:middle;border-bottom:none;padding-top:10px;padding-bottom:10px;">'

        $html += '<table id="legend" cellpadding="0" cellspacing="0" border="0" role="presentation" align="right" style="border-collapse:collapse;margin-left:auto;">'
        $html += '<tr>'

        $html += '<td valign="middle" style="vertical-align:middle;padding:0;border:1px solid #ccc;" width="' + $LegendLogoSize + '">'
        $html += '<img src="data:image/png;base64,' + $warningButtonBase64 + '" width="' + $LegendLogoSize + '" height="' + $LegendLogoSize + '" style="display:block;border:0;" alt="" />'
        $html += '</td>'
        $html += '<td valign="middle" style="vertical-align:middle;background-color:' + $barColors['Warning'] + ';color:#fff;padding:2px 6px;border:1px solid #ccc;line-height:' + $LegendLogoSize + 'px;">Warning</td>'

        $html += '<td valign="middle" style="vertical-align:middle;padding:0;border:1px solid #ccc;" width="' + $LegendLogoSize + '">'
        $html += '<img src="data:image/png;base64,' + $normalButtonBase64 + '" width="' + $LegendLogoSize + '" height="' + $LegendLogoSize + '" style="display:block;border:0;" alt="" />'
        $html += '</td>'
        $html += '<td valign="middle" style="vertical-align:middle;background-color:' + $barColors['Normal'] + ';color:#fff;padding:2px 6px;border:1px solid #ccc;line-height:' + $LegendLogoSize + 'px;">Normal</td>'

        $html += '<td valign="middle" style="vertical-align:middle;padding:0;border:1px solid #ccc;" width="' + $LegendLogoSize + '">'
        $html += '<img src="data:image/png;base64,' + $ignoreButtonBase64 + '" width="' + $LegendLogoSize + '" height="' + $LegendLogoSize + '" style="display:block;border:0;" alt="" />'
        $html += '</td>'
        $html += '<td valign="middle" style="vertical-align:middle;background-color:' + $barColors['Ignore'] + ';color:#fff;padding:2px 6px;border:1px solid #ccc;line-height:' + $LegendLogoSize + 'px;">Ignored</td>'

        $html += '</tr>'
        $html += '</table>'

        $html += '</td>'
        $html += '</tr>'
        $html += '</table>'

        # Main utilization table.
        $html += '<table id="tbl">'
        # $html += '<tr><td colspan="4"></td></tr>'
        $html += '<tr style="border-top: 2px solid #CCC;">'
        $html += '<td></td>'
        # $html += '<td width="420px">Name</td>'
        $html += '<td>Name</td>'
        # $html += '<td width="120px">Available</td>'
        $html += '<td>Available</td>'
        $html += '<td>&nbsp;&nbsp;Assigned / Total</td>'
        $html += '</tr>'

        # Data rows.
        # Sort by status first, then by available licenses so warnings with the lowest availability appear first.
        foreach ($item in $items | Sort-Object ThresholdStatusCode, Available) {
            $skuName = Get-HtmlEncodedText $item.SkuName

            $available = Format-HtmlNumber $item.Available
            $assigned = Format-HtmlNumber $item.Assigned
            $total = Format-HtmlNumber $item.Total

            # Calculate assigned-license usage ratio.
            $assignedValue = 0
            $totalValue = 0

            if ($null -ne $item.Assigned) {
                $assignedValue = [double]$item.Assigned
            }

            if ($null -ne $item.Total) {
                $totalValue = [double]$item.Total
            }

            $usedRatio = if ($totalValue -gt 0) {
                $assignedValue / $totalValue
            }
            else {
                0
            }

            # Keep the visual bar bounded between 0% and 100%.
            if ($usedRatio -lt 0) {
                $usedRatio = 0
            }

            if ($usedRatio -gt 1) {
                $usedRatio = 1
            }

            $filledWidth = [int][Math]::Round(($usedRatio * $BarWidth), 0)

            # Show a thin visible marker for very small non-zero usage.
            if (($assignedValue -gt 0) -and ($filledWidth -lt 1)) {
                $filledWidth = 1
            }

            $emptyWidth = $BarWidth - $filledWidth

            # Status-based color.
            $thresholdStatus = [string]$item.ThresholdStatus

            if ($barColors.ContainsKey($thresholdStatus)) {
                $barFillColor = Get-HtmlEncodedText $barColors[$thresholdStatus]
            }
            else {
                $barFillColor = Get-HtmlEncodedText $barColors['Ignore']
            }

            $barEmptyColor = '#CCC'

            # Build Outlook-safe bar as a nested table.
            # Avoid div/flex/percentage layouts for better Outlook compatibility.
            $barHtml = @()
            $barHtml += '<!-- Nested Table ' + $skuName + ' -->'
            $barHtml += '<table width="' + $BarWidth + '" cellpadding="0" cellspacing="0" border="0" role="presentation" style="border-collapse:collapse;border-spacing:0;table-layout:fixed;">'
            $barHtml += '<tr>'

            # Filled portion.
            if ($filledWidth -gt 0) {
                $barHtml += '<td width="' + $filledWidth + '" style="width:' + $filledWidth + 'px;height:' + $BarHeight + 'px;background-color:' + $barFillColor + ';padding:0;font-size:0;line-height:0;border-bottom:none;"></td>'
            }

            # Empty portion.
            if ($emptyWidth -gt 0) {
                $barHtml += '<td width="' + $emptyWidth + '" style="width:' + $emptyWidth + 'px;height:' + $BarHeight + 'px;background-color:' + $barEmptyColor + ';padding:0;font-size:0;line-height:0;border-bottom:none;"></td>'
            }

            $barHtml += '</tr>'
            $barHtml += '</table>'

            $barHtml = $barHtml -join ''

            # Assigned cell layout:
            # Nested table keeps the bar and assigned/total text aligned in Outlook.
            $assignedCell = @()

            # $assignedCell += '<table width="350" cellpadding="0" cellspacing="0" border="0" role="presentation" style="border-collapse:collapse;">'
            $assignedCell += '<table cellpadding="0" cellspacing="0" border="0" role="presentation" style="border-collapse:collapse;">'

            $assignedCell += '<tr>'
            $assignedCell += '<td width="' + $BarWidth + '" valign="middle" style="width:' + $BarWidth + 'px;border-bottom:none;vertical-align:middle;padding-right:0;">' + $barHtml + '</td>'

            # $assignedCell += '<td width="' + (350 - $BarWidth - 8) + '" valign="middle" style="white-space:nowrap;border-bottom:none;vertical-align:middle;">' + $assigned + ' / ' + $total + '</td>'
            $assignedCell += '<td valign="middle" style="white-space:nowrap;border-bottom:none;vertical-align:middle;">' + $assigned + ' / ' + $total + '</td>'
            $assignedCell += '</tr>'
            $assignedCell += '</table>'

            $assignedCell = $assignedCell -join ''

            $statusButtonBase64 = switch ($thresholdStatus) {
                'Normal' { $normalButtonBase64 }
                'Warning' { $warningButtonBase64 }
                'Ignore' { $ignoreButtonBase64 }
                default { $ignoreButtonBase64 }
            }

            $html += '<tr>'
            $html += '<td valign="middle" style="vertical-align:middle;padding-left:6px;" width="' + $LogoSize + '"><img src="data:image/png;base64,' + $statusButtonBase64 + '" width="' + $LogoSize + '" height="' + $LogoSize + '" style="display:block;" alt="" /></td>'
            $html += '<td valign="middle" style="vertical-align:middle;font-weight:bold;">' + $skuName + '</td>'
            $html += '<td valign="middle" style="vertical-align:middle;text-align:right;">' + $available + '</td>'
            $html += '<td valign="middle">' + $assignedCell + '</td>'
            $html += '</tr>'
        }

        # Footer spacer for main table.
        $html += '<tr><td class="head" colspan="4"></td></tr>'
        $html += '</table>'

        # Footer info.
        $module = Get-Module PsSkuMon365

        if ($module) {
            $moduleName = Get-HtmlEncodedText $module.Name
            $moduleVersion = Get-HtmlEncodedText $module.Version.ToString()
            $projectUri = Get-HtmlEncodedText $module.ProjectURI

            $html += '<table id="settings">'
            $html += '<tr><td colspan="2"><a href="' + $projectUri + '">' + $moduleName + ' v' + $moduleVersion + '</a></td></tr>'
            $html += '</table>'
        }

        $html += '</body></html>'

        return ($html -join "`n")
    }
}