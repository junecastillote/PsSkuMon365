function New-SkuMonList {
    [CmdletBinding()]
    param (

    )
    try {
        $subscribedSku = Get-MgSubscribedSku -ErrorAction Stop | Where-Object { $_.AppliesTo -eq 'User' -and $_.CapabilityStatus -eq 'Enabled' }
        $skuNames = Get-SkuFriendlyName -SkuPartNumber $($subscribedSku.SkuPartNumber) | Select-Object SkuName, SkuPartNumber
        $skuNames | Add-Member -MemberType NoteProperty -Name IncludeInReport -Value $True
        $skuNames | Add-Member -MemberType NoteProperty -Name AlertThreshold -Value 0

        return $skuNames
    }
    catch {
        throw "Failed to get subscribed sku list. $($_.Exception.Message)"
    }
}