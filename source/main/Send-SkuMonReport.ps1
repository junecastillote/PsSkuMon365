Function Send-SkuMonReport {
    [cmdletbinding()]
    param (
        [parameter(Mandatory)]
        [string]$From,

        [parameter()]
        [string[]]$To,

        [parameter()]
        [string[]]$CC,

        [parameter()]
        [string[]]$BCC,

        [parameter()]
        [string]$Subject,

        [parameter(Mandatory)]
        [string]$Body
    )

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

    if (!$To -and !$Cc -and !$Bcc) {
        SayError 'There must be at least 1 recipient.'
        retrn $null
    }

    $ThisFunction = ($MyInvocation.MyCommand)
    $ThisModule = Get-Module ($ThisFunction.Source)
    $ResourceFolder = [System.IO.Path]::Combine((Split-Path ($ThisModule.Path) -Parent), 'resource')
    $logo = $([convert]::ToBase64String([System.IO.File]::ReadAllBytes("$resourceFolder\logo.png")))

    # if ($PSVersionTable.PSEdition -eq 'Core') {
    #     $logo = $([convert]::ToBase64String((Get-Content $resourceFolder\logo.png -AsByteStream)))
    # }
    # else {
    #     $logo = $([convert]::ToBase64String((Get-Content $resourceFolder\logo.png -Raw -Encoding byte)))
    # }

    if (!$Subject) {
        $Subject = "m365 License Availability @ [$(($Body | Select-String -Pattern '(?<=\<b\>).+?(?=\<\/b\>)' -AllMatches).Matches[0].Value)]"
    }

    $mailBody = @{
        message = @{
            subject                = $Subject
            body                   = @{
                content     = $($Body.Replace("data:image/png;base64,$logo", "cid:logo"))
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