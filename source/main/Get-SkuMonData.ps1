function Get-SkuMonData {
    [CmdletBinding()]
    param (
        [Parameter(ValueFromPipeline)]
        $SkuMonList,

        [parameter()]
        [ValidateSet('PSObject', 'Html', 'Email')]
        [string]
        $OutputType = 'PSObject',

        [parameter()]
        [mailaddress]$From,

        [parameter()]
        [mailaddress[]]$To,

        [parameter()]
        [mailaddress[]]$CC,

        [parameter()]
        [mailaddress[]]$BCC
    )
    begin {

        if ($OutputType -in @('Html', 'Email')) {
            SayWarning "[$($MyInvocation.MyCommand.Name)]: The -OutputType parameter is being depracated and will be removed in future versions. Use ConvertTo-SkuMonHtml to generate the HTML output."
        }

        # If OutputType is Email, set the requirements.
        $err = 0
        if ($OutputType -eq 'Email') {
            SayWarning "[$($MyInvocation.MyCommand.Name)]: The -From,-TO,-CC,-BCC parameters are being depracated and will be removed in future versions. Use Send-SkuMonReport to send the report by email."
            if (!$From) {
                SayError "[$($MyInvocation.MyCommand.Name)]: The From email address is required."
                $err++
            }

            if (!$To -and !$Cc -and !$Bcc) {
                SayError "[$($MyInvocation.MyCommand.Name)]: There must be at least 1 recipient email address."
                $err++
            }
        }

        if ($err -gt 0) { continue }

        $thresholdStatusCode = @{
            Warning = 0
            Normal  = 1
            Ignore  = 2
        }

        if (!$SkuMonList) {
            $SkuMonList = New-SkuMonList
        }

        # $subscribedSku = Get-MgSubscribedSku -ErrorAction Stop | Where-Object { $_.AppliesTo -eq 'User' -and $_.CapabilityStatus -eq 'Enabled' }
        $subscribedSku = Get-MgSubscribedSku -ErrorAction Stop | Where-Object { $_.AppliesTo -eq 'User' }
        [System.Collections.ArrayList]$skuCollection = @()
    }

    process {
        # foreach ($item in $SkuMonList | Where-Object { $_.IncludeInReport -eq $true }) {
        foreach ($item in $SkuMonList) {
            $sku = $subscribedSku | Where-Object { $_.SkuPartNumber -eq $item.SkuPartNumber }

            $TotalUnits =
            $sku.PrepaidUnits.Enabled +
            $sku.PrepaidUnits.Warning

            $AvailableUnits = ($TotalUnits - $sku.ConsumedUnits)
            $ExcessUnits = 0
            if ($AvailableUnits -lt 0) {
                $ExcessUnits = [Math]::Abs($AvailableUnits)
                $AvailableUnits = 0
            }

            $thresholdStatus = $(
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

            $null = $skuCollection.Add(
                $(
                    New-Object psobject -Property (
                        [ordered]@{
                            PSTypeName          = 'SkuMonData'
                            SkuID               = $sku.SkuID
                            SkuPartNumber       = $sku.SkuPartNumber
                            SkuName             = $(
                                if (!($item.SkuName)) {
                                    $item.SkuPartNumber
                                }
                                else {
                                    $item.SkuName
                                }
                            )
                            Assigned            = $sku.ConsumedUnits
                            Total               = $TotalUnits
                            Enabled             = $sku.prepaidUnits.Enabled
                            Suspended           = $sku.prepaidUnits.Suspended
                            LockedOut           = $sku.PrepaidUnits.LockedOut
                            Warning             = $sku.prepaidUnits.Warning
                            Available           = $AvailableUnits
                            Invalid             = $ExcessUnits
                            CapabilityStatus    = $sku.CapabilityStatus
                            AlertThreshold      = $item.AlertThreshold
                            ThresholdStatus     = $thresholdStatus
                            ThresholdStatusCode = $thresholdStatusCode[$thresholdStatus]
                            ShowInReport        = $item.IncludeInReport
                        }
                    )
                )
            )
        }
    }

    end {

        if ($OutputType -eq 'Html' -or $OutputType -eq 'Email') {
            $html = $skuCollection | ConvertTo-SkuMonHtml
        }

        if ($OutputType -eq 'Html') {
            return $html
        }

        if ($OutputType -eq 'PSObject') {
            return $skuCollection
        }

        if ($OutputType -eq 'Email') {
            $params = @{
                From = $From
                Html = $html
            }

            if ($To.Count -gt 0) { $params.Add('To', $To) }
            if ($Cc.Count -gt 0) { $params.Add('Cc', $CC) }
            if ($BCC.Count -gt 0) { $params.Add('Bcc', $BCC) }

            try {
                Send-SkuMonReport @params -ErrorAction Stop
            }
            catch {
                throw "Failed to send email: $($_.Exception.Message)"
            }
        }
    }
}