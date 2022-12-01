Function Get-SkuMonData {
    [CmdletBinding()]
    param (
        [Parameter(ValueFromPipeline)]
        $SkuMonList,

        [parameter()]
        [ValidateSet('PSObject', 'Html', 'Email')]
        [string]
        $OutputType = 'PSObject',

        [parameter()]
        [string]$From,

        [parameter()]
        [string[]]$To,

        [parameter()]
        [string[]]$CC,

        [parameter()]
        [string[]]$BCC
    )
    begin {
        # If OutputType is Email, set the requirements.
        if ($OutputType -eq 'Email') {
            if (!$From) {
                Say 'The From email address is required.'
                return $null
            }

            if (!$To -and !$Cc -and !$Bcc) {
                Say 'There must be at least 1 recipient email address.'
                return $null
            }
        }

        # JSON email address conversion
        Function ConvertRecipientsToJSON {
            param(
                [Parameter(Mandatory)]
                [string[]]
                $Recipients
            )
            $jsonRecipients = @()
            $Recipients | ForEach-Object {
                $jsonRecipients += @{EmailAddress = @{Address = $_ } }
            }
            return $jsonRecipients
        }

        $ThisFunction = ($MyInvocation.MyCommand)
        $ThisModule = Get-Module ($ThisFunction.Source)

        if (!$SkuMonList) {
            $SkuMonList = New-SkuMonList
        }

        $subscribedSku = Get-MgSubscribedSku -ErrorAction Stop | Where-Object { $_.AppliesTo -eq 'User' }
        [System.Collections.ArrayList]$skuCollection = @()
    }

    process {
        foreach ($item in $SkuMonList | Where-Object { $_.IncludeInReport -eq $true }) {
            $sku = $subscribedSku | Where-Object { $_.SkuPartNumber -eq $item.SkuPartNumber }

            $AvailableUnits = (($sku.prepaidUnits.Enabled + $sku.prepaidUnits.Warning) - $sku.ConsumedUnits)
            $ExcessUnits = 0
            if ($AvailableUnits -lt 0) {
                $ExcessUnits = [Math]::Abs($AvailableUnits)
                $AvailableUnits = 0
            }

            $null = $skuCollection.Add(
                $(
                    New-Object psobject -Property (
                        [ordered]@{
                            PSTypeName       = 'SkuMonData'
                            SkuID            = $sku.SkuID
                            SkuPartNumber    = $sku.SkuPartNumber
                            SkuName          = $(
                                if (!($item.SkuName)) {
                                    $item.SkuPartNumber
                                }
                                else {
                                    $item.SkuName
                                }
                            )
                            Assigned         = $sku.ConsumedUnits
                            Total            = $sku.prepaidUnits.Enabled
                            Suspended        = $sku.prepaidUnits.Suspended
                            Warning          = $sku.prepaidUnits.Warning
                            Available        = $AvailableUnits
                            Invalid          = $ExcessUnits
                            CapabilityStatus = $sku.CapabilityStatus
                            AlertThreshold   = $item.AlertThreshold
                            ThresholdStatus  = $(
                                if ($item.AlertThreshold -gt 0) {
                                    if ($AvailableUnits -le $item.AlertThreshold) {
                                        "Warning"
                                    }

                                    if ($AvailableUnits -gt $item.AlertThreshold) {
                                        "Normal"
                                    }
                                }
                                else {
                                    "Ignore"
                                }
                            )
                        }
                    )
                )
            )
        }
    }

    end {

        if ($OutputType -eq 'Html' -or $OutputType -eq 'Email') {
            $Organization = Get-MgOrganization
            $ResourceFolder = [System.IO.Path]::Combine((Split-Path ($ThisModule.Path) -Parent), 'resource')
            $css = Get-Content $resourceFolder\style.css -Raw

            if ($PSVersionTable.PSEdition -eq 'Core') {
                $logo = $([convert]::ToBase64String((Get-Content $resourceFolder\logo.png -AsByteStream)))
            }
            else {
                $logo = $([convert]::ToBase64String((Get-Content $resourceFolder\logo.png -Raw -Encoding byte)))
            }

            $timeZoneInfo = [System.TimeZoneInfo]::Local
            $tz = $timeZoneInfo.DisplayName.ToString().Split(" ")[0]

            $today = Get-Date -Format g

            $title = 'Microsoft 365 License Availability Report'
            $html = @()
            $html += '<html><head><title>' + $title + '</title>'
            $html += '<style type="text/css">'
            $html += $css
            $html += '</style></head>'
            $html += '<body>'
            #table headers
            $html += '<table id="tbl">'
            $html += '<tr><td class="head"> </td></tr>'
            $html += '<tr><th class="section">Microsoft 365 Licenses</th></tr>'
            $html += '<tr><td class="head"><b>' + $($Organization.DisplayName) + '</b><br>' + $today + ' ' + $tz + '</td></tr>'
            $html += '<tr><td class="head"> </td></tr>'
            # $html += '<tr><td class="head"> </td></tr>'
            $html += '</table>'
            $html += '<tr><td class="head" colspan="4"></td></tr>'
            $html += '</table>'
            $html += '<table id="legend">'
            $html += '<tr><td class="green" width="60px">Normal</td><td class="red" width="60px">Warning</td><td class="gray" width="60px">Ignore</td></tr>'
            $html += '</table>'
            $html += '<table id="tbl">'
            $html += '<tr><td width="420px" colspan="2">Name</th><td width="170px">Quantity</td><td width="5px"></td></tr>'

            foreach ($item in ($skuCollection | Sort-Object ThresholdStatus -Descending)) {
                $html += '<tr><td><img src="data:image/png;base64,' + $logo + '"></img></td>'
                $html += '<th>' + $item.SkuName + '</th>'
                $html += '<td><b>' + $('{0:N0}' -f $item.Available) + ' available</b><br>' + $('{0:N0}' -f $item.Assigned) + ' assigned out of ' + $('{0:N0}' -f $item.Total) #+ ' total'

                if (($item.ThresholdStatus) -eq 'Warning') {
                    $html += '<td class="red" width="5px"></td></tr>'
                }
                if (($item.ThresholdStatus) -eq 'Normal') {
                    $html += '<td class="green" width="5px"></td></tr>'
                }
                if (($item.ThresholdStatus) -eq 'Ignore') {
                    $html += '<td class="gray" width="5px"></td></tr>'
                }
            }

            $html += '<tr><td class="head" colspan="4"></td></tr>'
            $html += '</table>'

            $html += '<table id="legend">'
            $html += '<tr><td class="green" width="60px">Normal</td><td class="red" width="60px">Warning</td><td class="gray" width="60px">Ignore</td></tr>'
            $html += '</table>'
            $html += '<table id="settings">'
            $html += '<tr><td colspan="2"><a href="' + $ThisModule.ProjectURI + '">' + $ThisModule.Name + ' v' + $ThisModule.Version + '</a></td></tr>'
            $html += '</table>'
            $html += '</body>'
            $html += '</html>'
            $html = ($html -join "`n")
        }

        if ($OutputType -eq 'Html') {
            $html
        }

        if ($OutputType -eq 'PSObject') {
            $skuCollection
        }

        if ($OutputType -eq 'Email') {
            $ResourceFolder = [System.IO.Path]::Combine((Split-Path ($ThisModule.Path) -Parent), 'resource')
            $logo = $([convert]::ToBase64String([System.IO.File]::ReadAllBytes("$resourceFolder\logo.png")))
            $Subject = "[$($Organization.DisplayName)] Microsoft 365 License Availability"
            $mailBody = @{
                message = @{
                    subject                = $Subject
                    body                   = @{
                        content     = $($html.Replace("data:image/png;base64,$logo", "cid:logo"))
                        contentType = "HTML"
                    }
                    internetMessageHeaders = @(
                        @{
                            name  = "X-Mailer"
                            value = "PsGraphMail by june.castillote@gmail.com"
                        }
                    )
                    attachments            = @(
                        @{
                            "@odata.type"  = "#microsoft.graph.fileAttachment"
                            "contentID"    = "logo"
                            "name"         = "logo"
                            "IsInline"     = $true
                            "contentType"  = "image/png"
                            "contentBytes" = $logo
                        }
                    )
                }
            }

            # To recipients
            if ($To) {
                $mailBody.message += @{
                    toRecipients = @(
                        $(ConvertRecipientsToJSON $To)
                    )
                }
            }

            # Cc recipients
            if ($CC) {
                $mailBody.message += @{
                    ccRecipients = @(
                        $(ConvertRecipientsToJSON $CC)
                    )
                }
            }

            # BCC recipients
            if ($BCC) {
                $mailBody.message += @{
                    bccRecipients = @(
                        $(ConvertRecipientsToJSON $BCC)
                    )
                }
            }

            try {
                Send-MgUserMail -UserId $From -BodyParameter $mailBody
            }
            catch {
                SayError "Send email failed: $($_.Exception.Message)"
            }
        }
    }
}