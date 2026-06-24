function Invoke-SkuMon {
    [CmdletBinding()]
    param (
        # Input configuration
        [Parameter(ValueFromPipeline)]
        [object[]]$SkuMonList,

        # Output control
        [ValidateSet('Object', 'Html', 'Email')]
        [string]$Mode = 'Object',

        # HTML options
        [string]$ReportTitle = 'Microsoft 365 License Availability Report',
        [string]$OrganizationName,
        [bool]$ShowLegend = $true,

        # Email options
        [mailaddress]$From,
        [mailaddress[]]$To,
        [mailaddress[]]$Cc,
        [mailaddress[]]$Bcc,
        [string]$Subject,

        [switch]$IncludeCsv
    )

    begin {

        $items = [System.Collections.Generic.List[object]]::new()

        if ($Mode -eq 'Email') {
            if (-not $From) {
                throw "The -From parameter is required when Mode = Email."
            }
            if (-not ($To -or $Cc -or $Bcc)) {
                throw "At least one recipient (-To, -Cc, -Bcc) is required."
            }
        }

    }

    process {

        if ($SkuMonList) {
            foreach ($item in $SkuMonList) {
                $items.Add($item)
            }
        }

    }

    end {

        # 1. Resolve input list
        if ($items.Count -eq 0) {
            Write-Verbose "[$($MyInvocation.MyCommand.Name)]: No input list provided. Using New-SkuMonList..."
            $items = New-SkuMonList
        }

        # 2. Get data
        Write-Verbose "[$($MyInvocation.MyCommand.Name)]: Collecting SKU monitoring data..."
        $data = $items | Get-SkuMonData -ErrorAction Stop

        # 3. Return PSObject (default / fast path)
        if ($Mode -eq 'Object') {
            return $data
        }

        # 4. Convert to HTML
        Write-Verbose "[$($MyInvocation.MyCommand.Name)]: Converting data to HTML..."
        $html = $data | ConvertTo-SkuMonHtml `
            -ReportTitle $ReportTitle `
            -OrganizationName $OrganizationName `
            -ShowLegend $ShowLegend

        if ($Mode -eq 'Html') {
            return $html
        }

        # 5. Email mode
        if ($Mode -eq 'Email') {
            Write-Verbose "[$($MyInvocation.MyCommand.Name)]: Sending report via email..."

            $params = @{
                From = $From
                Html = $html
            }

            if ($To) { $params.To = $To }
            if ($Cc) { $params.Cc = $Cc }
            if ($Bcc) { $params.Bcc = $Bcc }
            if ($Subject) { $params.Subject = $Subject }
            if ($OrganizationName) { $params.OrganizationName = $OrganizationName }
            if ($IncludeCsv) { $params.CsvObject = ($data | Select-Object -ExcludeProperty ThresholdStatusCode) }

            Send-SkuMonReport @params -ErrorAction Stop

            # Return useful output even in email mode
            return [pscustomobject]@{
                Status     = 'Sent'
                Recipients = ($To + $Cc + $Bcc)
                Timestamp  = (Get-Date)
            }
        }
    }
}