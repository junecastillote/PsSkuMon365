#Dot-Source all functions
Get-ChildItem -Path $PSScriptRoot\source\*.ps1 -Recurse |
ForEach-Object {
    . $_.FullName
}