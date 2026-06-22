function Send-MailMessageByGraph {
    [CmdletBinding()]
    param (
        [parameter(Mandatory)]
        [string]$From,

        [parameter()]
        [string[]]$To,

        [parameter()]
        [string[]]$CC,

        [parameter()]
        [string[]]$BCC,

        [parameter(Mandatory)]
        [string]$Subject,

        [parameter(Mandatory)]
        [string]$Body,

        [parameter()]
        [string[]]$Attachment,

        [parameter()]
        [ValidateSet('Low', 'Normal', 'High')]
        [string]$Importance = 'Normal'
    )

    # Validate that there is at least one recipient before building the Graph payload.
    # Graph mail requires a recipient list; failing early prevents a malformed API call.
    if (!$To -and !$CC -and !$BCC) {
        throw "At least one To, Cc, or Bcc recipient is required."
    }

    function ConvertRecipientsToJSON {
        param(
            [Parameter(Mandatory)]
            [string[]]
            $Recipients
        )

        # Convert recipient email addresses into the Graph JSON shape.
        # Keeping this conversion separate makes recipient handling reusable
        # for To/Cc/Bcc and improves readability of the main payload builder.
        $jsonRecipients = @()
        $Recipients | ForEach-Object {
            $jsonRecipients += @{EmailAddress = @{Address = $_ } }
        }
        return $jsonRecipients
    }

    # Build the base mail payload for Graph API submission.
    # This payload is intentionally minimal and extended only with actual
    # recipients/attachments present in the request.
    $mailBody = @{
        message = @{
            subject                = $Subject
            importance             = $Importance
            body                   = @{
                content     = $Body
                contentType = "HTML"
            }
            internetMessageHeaders = @(
                @{
                    name  = "X-Mailer"
                    value = "PsGraphMail by june.castillote@gmail.com"
                }
            )
            attachments            = @()
        }
    }

    # Add recipient blocks only when the corresponding parameter is supplied.
    # This avoids sending empty recipient arrays and keeps the request lean.
    if ($To) {
        $mailBody.message += @{
            toRecipients = @(
                $(ConvertRecipientsToJSON $To)
            )
        }
    }

    if ($CC) {
        $mailBody.message += @{
            ccRecipients = @(
                $(ConvertRecipientsToJSON $CC)
            )
        }
    }

    if ($BCC) {
        $mailBody.message += @{
            bccRecipients = @(
                $(ConvertRecipientsToJSON $BCC)
            )
        }
    }

    # Attach files to the message when attachments are provided.
    # Each file is read and base64-encoded to satisfy Graph file attachment requirements.
    if ($Attachment) {
        foreach ($file in $Attachment) {
            try {
                $filePath = (Resolve-Path $file -ErrorAction STOP).Path

                $fileName = $(Split-Path $filePath -Leaf)

                $fileByte = "$([convert]::ToBase64String([System.IO.File]::ReadAllBytes($filePath)))"

                $mailBody.message.attachments += @{
                    "@odata.type"  = "#microsoft.graph.fileAttachment"
                    "name"         = $fileName
                    "contentBytes" = $fileByte
                }
            }
            catch {
                throw "Attachment failed ($($fileName)): $($_.Exception.Message)"
            }
        }
    }

    # Submit the assembled payload to Microsoft Graph.
    # Any failure here is surfaced explicitly so callers can handle send problems.
    try {
        Send-MgUserMail -UserId $From -BodyParameter $mailBody -ErrorAction Stop
    }
    catch {
        throw "Send email failed: $($_.Exception.Message)"
    }
}