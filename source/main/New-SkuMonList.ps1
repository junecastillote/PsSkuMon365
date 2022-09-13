Function New-SkuMonList {
    [CmdletBinding()]
    param (

    )
    try {
        $subscribedSku = Get-MgSubscribedSku -ErrorAction Stop | Where-Object { $_.AppliesTo -eq 'User' }
        $skuNames = Get-SkuFriendlyName -SkuPartNumber $($subscribedSku.SkuPartNumber) | Select-Object SkuName, SkuPartNumber
        $skuNames | Add-Member -MemberType NoteProperty -Name IncludeInReport -Value $True
        $skuNames | Add-Member -MemberType NoteProperty -Name AlertThreshold -Value 0

        return $skuNames
    }
    catch {
        SayError $_.Exception.Message
    }
}