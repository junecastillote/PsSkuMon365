Function Get-SkuFriendlyName {
    [CmdletBinding(DefaultParameterSetName = 'All')]
    param (
        ## This is URL path to the the licensing reference table document from GitHub.
        ## The current working URL is the default value.
        ## In case Microsoft moved the document, use this parameter to point to the new URL.

        [parameter()]
        [string]
        $URL = 'https://raw.githubusercontent.com/MicrosoftDocs/azure-docs/master/articles/active-directory/enterprise-users/licensing-service-plan-reference.md',

        [parameter()]
        [string[]]
        $SkuPartNumber,

        ## Convert license names to title case.
        [parameter()]
        [switch]
        $TitleCase
    )

    $ErrorActionPreference = 'STOP'

    #https://docs.microsoft.com/en-us/azure/active-directory/enterprise-users/licensing-service-plan-reference

    ## Parse the Markdown Table from the $URL
    try {
        [System.Collections.ArrayList]$raw_Table = ([System.Net.WebClient]::new()).DownloadString($URL).split("`n")
    }
    catch {
        SayError "There was an error getting the licensing reference table at [$URL]. Please make sure that the URL is still valid."
        SayError $_.Exception.Message
        return $null
    }

    ## Determine the starting row index of the table
    $startLine = $raw_Table.IndexOf('| Product name | String ID | GUID | Service plans included | Service plans included (friendly names) |')

    ## Determine the ending index of the table
    $endLine = ($raw_Table.IndexOf('## Service plans that cannot be assigned at the same time') - 1)

    ## Extract the string in between the lines $startLine and $endLine
    $result = for ($i = $startLine + 1; $i -lt $endLine; $i++) {
        if ($raw_Table[$i] -notlike "*---*") {
            $raw_Table[$i].Substring(1, $raw_Table[$i].Length - 1)
        }
    }

    ## Perform a little clean-up
    ### replace "[space] | [space]" with "|"
    ### replace "[space]<br/>[space]" with ","
    ### replace "((" with "("
    ### replace "))" with ")"
    ### #replace ")[space](" with ")("

    $result = $result `
        -replace '\s*\|\s*', '|' `
        -replace '\s*<br/>\s*', ',' `
        -replace '\(\(', '(' `
        -replace '\)\)', ')' `
        -replace '\)\s*\(', ')('

    ## Create the result object
    $result = @($result | ConvertFrom-Csv -Delimiter "|" -Header 'SkuName', 'SkuPartNumber', 'SkuID', 'ChildServicePlan', 'ChildServicePlanName')

    if ($TitleCase) {

        ## Convert product name to title case
        $TextInfo = (Get-Culture).TextInfo
        for ($i = 0; $i -lt $result.Count; $i++) {
            $result[$i].SkuName = $TextInfo.ToTitleCase(($result[$i].SkuName).ToLower())
            $result[$i].ChildServicePlanName = $TextInfo.ToTitleCase(($result[$i].ChildServicePlanName).ToLower())
        }
    }

    if ($SkuPartNumber) {
        [System.Collections.ArrayList]$filteredResult = @()
        foreach ($sku in $SkuPartNumber) {
            $item = ($result | Where-Object { $_.SkuPartNumber -eq $sku })
            if ($item) {
                $null = $filteredResult.Add($item)
            }
            if (!$item) {
                $null = $filteredResult.Add($( New-Object psobject -Property ([ordered]@{
                            SkuName              = $null
                            SkuPartNumber        = $sku
                            SkuId                = $null
                            ChildServicePlan     = $null
                            ChildServicePlanName = $null
                        })
                    ))
            }
        }
        return $filteredResult
    }

    if (!$SkuPartNumber) {
        return $result
    }
}
