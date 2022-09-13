[CmdletBinding()]
param (
    [string]$ModulePath
)
$moduleManifest = Test-ModuleManifest -Path $((Get-ChildItem -Path $PSScriptRoot -Filter *.psd1).FullName)
$module = Get-Module $($moduleManifest.Name) -ListAvailable #| Where-Object { $_.Version -eq $moduleManifest.Version })

if ($module) {
    foreach ($Moduleinfo in $module) {
        $ModulePath = $Moduleinfo.ModuleBase
        try {
            $items = Get-ChildItem $ModulePath -Recurse

            # Delete all module files
            foreach ($file in (($items | Where-Object {!$_.PSIsContainer}).FullName)) {
                # $file
                [System.IO.File]::Delete($file)
            }

            # Delete all module subfolders
            $folders = @(($items | Where-Object {$_.PSIsContainer}).FullName)
            [array]::Reverse($folders)
            foreach ($folder in $folders) {
                [System.IO.Directory]::Delete($folder)
            }

            # Delete module folder and parent folder
            [System.IO.Directory]::Delete($ModulePath)
            [System.IO.Directory]::Delete($(Split-Path $ModulePath))

            # Remove-Item -Path $ModulePath -Recurse -Force -Confirm:$false -ErrorAction Stop
            Write-Output "Done uninstalling $($Moduleinfo.Name) version $($Moduleinfo.Version) from $($Moduleinfo.ModuleBase)"
        }
        catch {
            Write-Output ""
            Write-Output "Failed"
            Write-Output $_.Exception.Message
            Write-Output ""
            return $null
        }
    }
}
else {
    "[$($moduleManifest.Name)] module not found. Nothing to uninstall." | Out-Default
}