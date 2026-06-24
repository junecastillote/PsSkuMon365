function Send-SkuMonReport {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory)]
        [string]$From,

        [Parameter()]
        [mailaddress[]]$To,

        [Parameter()]
        [mailaddress[]]$Cc,

        [Parameter()]
        [mailaddress[]]$Bcc,

        [Parameter(Mandatory)]
        [string]$Html,

        [Parameter()]
        [string]$Subject,

        [Parameter()]
        [string]$OrganizationName,

        [Parameter()]
        [psobject[]]$CsvObject,

        [Parameter()]
        [string]$CsvString
    )

    # Resolve module resource folder (for logo)
    $module = Get-Module PsSkuMon365

    if (-not $module) {
        throw "Module 'PsSkuMon365' is not loaded."
    }

    $ResourceFolder = Join-Path (Split-Path $module.Path -Parent) 'resource'
    # Load PNG as base64
    $normalButtonPath = Join-Path $ResourceFolder 'normal.png'
    $ignoreButtonPath = Join-Path $ResourceFolder 'ignore.png'
    $warningButtonPath = Join-Path $ResourceFolder 'warning.png'

    $normalButtonBase64 = "$([convert]::ToBase64String([System.IO.File]::ReadAllBytes($normalButtonPath)))"
    $ignoreButtonBase64 = "$([convert]::ToBase64String([System.IO.File]::ReadAllBytes($ignoreButtonPath)))"
    $warningButtonBase64 = "$([convert]::ToBase64String([System.IO.File]::ReadAllBytes($warningButtonPath)))"

    # Organization info
    if (-not $OrganizationName) {
        $org = Get-MgOrganization -ErrorAction Stop
        $OrganizationName = $org.DisplayName
    }
    else {
        $OrganizationName = [System.Net.WebUtility]::HtmlEncode($OrganizationName)
    }

    # Get org name for subject if not supplied
    if (-not $Subject) {
        $Subject = "[$($OrganizationName)] Microsoft 365 License Availability"
    }

    # IMPORTANT:
    # Replace base64 logo references with CID
    # (current HTML uses data:image/png;base64,...)

    $HtmlBody = $Html.Replace(
        'data:image/png;base64,' + $normalButtonBase64, 'cid:normal').Replace(
        'data:image/png;base64,' + $warningButtonBase64, 'cid:warning').Replace(
        'data:image/png;base64,' + $ignoreButtonBase64, 'cid:ignore')

    # Helper: Convert recipients
    function ConvertTo-GraphRecipients {
        param ($Addresses)
        $Addresses | ForEach-Object {
            @{
                EmailAddress = @{
                    Address = $_
                }
            }
        }
    }

    # CSV Attachment handling
    $csvAttachment = $null
    $CsvFileName = "SkuMonReport.csv"

    if ($CsvObject -or $CsvString) {
        Write-Verbose "[$($MyInvocation.MyCommand.Name)]: Adding CSV attachment..."
        try {
            if ($CsvObject) {
                # Convert object to CSV string (no type info, Outlook/Excel friendly)
                $CsvString = $CsvObject | ConvertTo-Csv -NoTypeInformation | Out-String
            }

            # Normalize line endings (important for Excel)
            $CsvString = $CsvString -replace "`r?`n", "`r`n"

            # Encode to Base64
            $csvBytes = [System.Text.Encoding]::UTF8.GetBytes($CsvString)
            $csvBase64 = [convert]::ToBase64String($csvBytes)

            # Build Graph attachment object
            $csvAttachment = @{
                "@odata.type" = "#microsoft.graph.fileAttachment"
                name          = $CsvFileName
                contentType   = "text/csv"
                contentBytes  = $csvBase64
            }
        }
        catch {
            throw "Failed to build CSV attachment: $($_.Exception.Message)"
        }
    }

    $attachments = @(
        @{
            "@odata.type" = "#microsoft.graph.fileAttachment"
            name          = "normal.png"
            contentId     = "normal"
            isInline      = $true
            contentType   = "image/png"
            contentBytes  = $normalButtonBase64
        },
        @{
            "@odata.type" = "#microsoft.graph.fileAttachment"
            name          = "warning.png"
            contentId     = "warning"
            isInline      = $true
            contentType   = "image/png"
            contentBytes  = $warningButtonBase64
        },
        @{
            "@odata.type" = "#microsoft.graph.fileAttachment"
            name          = "ignore.png"
            contentId     = "ignore"
            isInline      = $true
            contentType   = "image/png"
            contentBytes  = $ignoreButtonBase64
        }
    )

    # Write-Verbose "Hello 123"
    # Add CSV if present
    if ($csvAttachment) {
        $attachments += $csvAttachment
    }

    # Build message
    $mailBody = @{
        message = @{
            subject     = $Subject
            body        = @{
                contentType = "HTML"
                content     = $HtmlBody
            }
            attachments = $attachments
        }
    }

    if ($To) {
        $mailBody.message.toRecipients = @(ConvertTo-GraphRecipients $To)
    }

    if ($Cc) {
        $mailBody.message.ccRecipients = @(ConvertTo-GraphRecipients $Cc)
    }

    if ($Bcc) {
        $mailBody.message.bccRecipients = @(ConvertTo-GraphRecipients $Bcc)
    }

    try {
        Send-MgUserMail -UserId $From -BodyParameter $mailBody -ErrorAction Stop
    }
    catch {
        throw "Failed to send email: $($_.Exception.Message)"
    }
}