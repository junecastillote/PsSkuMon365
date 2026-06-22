function ConvertTo-SkuMonHtml {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory, ValueFromPipeline)]
        [psobject[]]$InputObject,

        [Parameter()]
        [string]$ReportTitle = 'Microsoft 365 License Availability Report'
    )

    begin {
        $items = [System.Collections.Generic.List[object]]::new()
        $LogoSize = 24
        $BarWidth = 100
        $BarHeight = 8
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

        # Resolve resource folder
        # if (-not $ResourceFolder) {
        $module = Get-Module PsSkuMon365

        if (-not $module) {
            throw "Module 'PsSkuMon365' is not loaded. Unable to resolve the resource folder."
        }

        $ResourceFolder = Join-Path (Split-Path $module.Path -Parent) 'resource'
        # }

        # Load CSS
        $cssPath = Join-Path $ResourceFolder 'style.css'
        $css = Get-Content $cssPath -Raw

        # Load logo as base64
        $logoPath = Join-Path $ResourceFolder 'logo.png'

        if ($PSVersionTable.PSEdition -eq 'Core') {
            $logoBytes = Get-Content $logoPath -AsByteStream
        }
        else {
            $logoBytes = Get-Content $logoPath -Encoding Byte -Raw
        }

        $logoBase64 = [Convert]::ToBase64String($logoBytes)

        # Organization info
        $org = Get-MgOrganization -ErrorAction Stop
        $orgName = [System.Net.WebUtility]::HtmlEncode($org.DisplayName)

        # Time info
        $tzInfo = [System.TimeZoneInfo]::Local
        $tz = ($tzInfo.DisplayName -split ' ')[0]
        $today = Get-Date -Format g

        $html = @()

        # Header
        $html += '<html><head><title>' + [System.Net.WebUtility]::HtmlEncode($ReportTitle) + '</title>'
        $html += '<style type="text/css">'
        $html += $css
        $html += '</style></head><body>'

        # Title section
        $html += '<table id="tbl">'
        $html += '<tr><td class="head"></td></tr>'
        $html += '<tr><th class="section">Microsoft 365 Licenses</th></tr>'
        $html += '<tr><td class="head"><b>' + $orgName + '</b><br>' + $today + ' ' + $tz + '</td></tr>'
        $html += '<tr><td class="head"></td></tr>'
        $html += '</table>'

        # Legend
        $html += '<table id="legend">'
        $html += '<tr>'
        $html += '<td class="Normal" width="60px">Normal</td>'
        $html += '<td class="Warning" width="60px">Warning</td>'
        $html += '<td class="Ignore" width="60px">Ignore</td>'
        $html += '</tr>'
        $html += '</table>'

        # Table header
        $html += '<table id="tbl">'
        $html += '<tr>'
        $html += '<td width="420px" colspan="2">Name</td>'
        $html += '<td width="120px">Available licenses</td>'
        $html += '<td width="190px">Assigned licenses</td>'
        $html += '<td width="5px"></td>'
        $html += '</tr>'

        # Data rows
        foreach ($item in $items | Sort-Object ThresholdStatus, Available -Descending) {

            $skuName = [System.Net.WebUtility]::HtmlEncode($item.SkuName)

            $available = '{0:N0}' -f $item.Available
            $assigned = '{0:N0}' -f $item.Assigned
            $total = '{0:N0}' -f $item.Total

            $statusClass = [System.Net.WebUtility]::HtmlEncode($item.ThresholdStatus)

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
            # Normal uses your warmer muted tone.
            $barFillColor = switch ($item.ThresholdStatus) {
                'Warning' { '#C0392B' }
                'Normal' { '#8A5A00' }
                'Ignore' { '#7F8C8D' }
                default { '#8A5A00' }
            }

            $barEmptyColor = '#D9D9D9'

            # Build Outlook-safe bar as a nested table.
            # Avoid div/flex/percentage layouts for better Outlook compatibility.
            $barHtml = @()
            $barHtml += "<!-- Nested Table $SkuName -->"
            $barHtml += '<table width="' + $BarWidth + '" cellpadding="0" cellspacing="0" border="0" role="presentation" style="border-collapse:collapse;table-layout:fixed;">'
            $barHtml += '<tr>'

            # Filled portion
            if ($filledWidth -gt 0) {
                $barHtml += "<!-- Filled portion $SkuName -->"
                $barHtml += '<td width="' + $filledWidth + '" style="width:' + $filledWidth + 'px;padding:0;" valign="middle">'
                $barHtml += '<table width="' + $filledWidth + '" height="' + $BarHeight + '" cellpadding="0" cellspacing="0" style="border-collapse:collapse;">'
                $barHtml += '<tr>'
                $barHtml += '<td bgcolor="' + $barFillColor + '" style="font-size:0;line-height:0;border-bottom:none;"></td>'
                $barHtml += '</tr></table>'
                $barHtml += '</td>'
            }

            # Empty portion
            if ($emptyWidth -gt 0) {
                $barHtml += "<!-- Empty portion $SkuName -->"
                $barHtml += '<td width="' + $emptyWidth + '" style="width:' + $emptyWidth + 'px;padding:0;" valign="middle">'
                $barHtml += '<table width="' + $emptyWidth + '" height="' + $BarHeight + '" cellpadding="0" cellspacing="0" style="border-collapse:collapse;">'
                $barHtml += '<tr>'
                $barHtml += '<td bgcolor="' + $barEmptyColor + '" style="font-size:0;line-height:0;border-bottom:none;"></td>'
                $barHtml += '</tr></table>'
                $barHtml += '</td>'
            }

            $barHtml += '</tr>'
            $barHtml += '</table>'

            $barHtml = $barHtml -join ''

            # Assigned cell layout:
            # Nested table keeps the bar and assigned/total text aligned in Outlook.

            $assignedCell = @()
            $assignedCell += '<table width="190" cellpadding="0" cellspacing="0" border="0" role="presentation" style="border-collapse:collapse;">'
            $assignedCell += '<tr>'
            $assignedCell += '<td width="' + $BarWidth + '" valign="middle" style="width:' + $BarWidth + 'px;border-bottom:none;vertical-align:middle;">' + $barHtml + '</td>'
            $assignedCell += '<td width="8" style="width:8px;font-size:0;line-height:0;border-bottom:none;vertical-align:middle;">&nbsp;</td>'
            $assignedCell += '<td width="' + (190 - $BarWidth - 8) + '" valign="middle" style="white-space:nowrap;border-bottom:none;vertical-align:middle;">' + $assigned + '/' + $total + '</td>'
            $assignedCell += '</tr>'
            $assignedCell += '</table>'
            $assignedCell = $assignedCell -join ''

            $html += '<tr>'
            $html += '<td valign="middle" style="vertical-align:middle;"><img src="data:image/png;base64,' + $logoBase64 + '" width="' + $LogoSize + '" height="' + $LogoSize + '" style="display:block;" alt="" /></td>'
            $html += '<td valign="middle" style="vertical-align:middle;font-weight: bold;">' + $skuName + '</td>'
            $html += '<td valign="middle" style="vertical-align:middle;">' + $available + '</td>'
            $html += '<td valign="middle">' + $assignedCell + '</td>'
            $html += '<td class="' + $statusClass + '" width="5px" valign="middle"></td>'
            $html += '</tr>'
        }

        # Footer spacer
        $html += '<tr><td class="head" colspan="5"></td></tr>'
        $html += '</table>'

        # Legend (bottom)
        $html += '<table id="legend">'
        $html += '<tr>'
        $html += '<td class="Normal" width="60px">Normal</td>'
        $html += '<td class="Warning" width="60px">Warning</td>'
        $html += '<td class="Ignore" width="60px">Ignore</td>'
        $html += '</tr>'
        $html += '</table>'

        # Footer info
        $module = Get-Module PsSkuMon365

        if ($module) {
            $moduleName = [System.Net.WebUtility]::HtmlEncode($module.Name)
            $moduleVersion = [System.Net.WebUtility]::HtmlEncode($module.Version.ToString())
            $projectUri = [System.Net.WebUtility]::HtmlEncode($module.ProjectURI)

            $html += '<table id="settings">'
            $html += '<tr><td colspan="2"><a href="' + $projectUri + '">' + $moduleName + ' v' + $moduleVersion + '</a></td></tr>'
            $html += '</table>'
        }

        $html += '</body></html>'

        return ($html -join "`n")
    }
}