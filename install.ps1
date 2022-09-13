[CmdletBinding()]
param (
    [parameter()]
    [ValidateSet('CurrentUser', 'AllUsers')]
    $Scope
)

$moduleManifest = Get-ChildItem -Path $PSScriptRoot -Filter *.psd1
$Moduleinfo = Test-ModuleManifest -Path ($moduleManifest.FullName)

Remove-Module ($Moduleinfo.Name) -ErrorAction SilentlyContinue


$isWindowsOs = ([System.Environment]::OSVersion.Platform -eq 'Win32NT')

# Is elevated PS?
if ($isWindowsOs) {
    ## On Windows
    [bool]$isAdmin = ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
}
else {
    ## On non-Windows
    [bool]$isAdmin = ([System.Environment]::UserName -eq 'root')
}

# If -Scope is not specified, set the scope based on whether PowerShell is running as admin / root or a regular user.
if (!$Scope -and $isAdmin) { $Scope = 'AllUsers' }
if (!$Scope -and !$isAdmin) { $Scope = 'CurrentUser' }

# Do not install if the PS session is not root or admin and scope is all users.
if ($Scope -eq 'AllUsers' -and !$isAdmin ) {
    "Installing the module in AllUsers scope requires an elevated PowerShell (run as admin or sudo)." | Out-Default
    return $null
}

# Windows + PowerShell Core + CurrentUser
if ($PSEdition -eq 'Core' -and $isWindowsOs -and $Scope -eq 'CurrentUser') {
    $ModulePath = ([System.IO.Path]::Combine(([Environment]::GetFolderPath("MyDocuments")), 'PowerShell', 'Modules'))
}

# Windows + PowerShell Core + AllUsers
if ($PSEdition -eq 'Core' -and $isWindowsOs -and $Scope -eq 'AllUsers') {
    $ModulePath = "$env:ProgramFiles\PowerShell\Modules"
}

# Windows PowerShell + CurrentUser
if ($PSEdition -eq 'Desktop' -and $Scope -eq 'CurrentUser') {
    $ModulePath = ([System.IO.Path]::Combine(([Environment]::GetFolderPath("MyDocuments")), 'WindowsPowerShell', 'Modules'))
}

# Windows PowerShell + AllUsers
if ($PSEdition -eq 'Desktop' -and $Scope -eq 'AllUsers') {
    $ModulePath = "$env:ProgramFiles\WindowsPowerShell\Modules"
}

# Non-Windows + CurrentUser
if (!$isWindowsOs -and $Scope -eq 'CurrentUser') {
    $ModulePath = "$HOME/.local/share/powershell/Modules"
}
# Non-Windows + AllUsers
if (!$isWindowsOs -and $Scope -eq 'AllUsers') {
    $ModulePath = '/usr/local/share/powershell/Modules'
}

$ModulePath = $ModulePath + "/$($Moduleinfo.Name.ToString())/$($Moduleinfo.Version.ToString())"
$ModulePath

if (!(Test-Path $ModulePath)) {
    New-Item -Path $ModulePath -ItemType Directory -Force | Out-Null
}

try {
    Copy-Item -Path $PSScriptRoot\* -Include *.psd1, *.psm1 -Destination $ModulePath -Force -Confirm:$false -ErrorAction Stop
    Copy-Item -Path $PSScriptRoot\source -Recurse -Destination $ModulePath -Force -Confirm:$false -ErrorAction Stop
    Copy-Item -Path $PSScriptRoot\resource -Recurse -Destination $ModulePath -Force -Confirm:$false -ErrorAction Stop
    Write-Output ""
    Write-Output "Success. Installed to $ModulePath"
    Write-Output ""
    if ($isWindowsOs) {
        Get-ChildItem -Recurse $ModulePath | Unblock-File -Confirm:$false
    }
    Import-Module $($Moduleinfo.Name.ToString()) -Force
}
catch {
    Write-Output ""
    Write-Output "Failed"
    Write-Output $_
    Write-Output ""
}