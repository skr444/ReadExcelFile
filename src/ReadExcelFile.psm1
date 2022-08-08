<#
    .SYNOPSIS
    ReadExcelFile Powershell module main file.
 
    .DESCRIPTION
    Module entry point containing the logic that is executed when
    the module is imported into the session with Import-Module.
#>

#Requires -Version 5
Set-StrictMode -Version 3
$ErrorActionPreference = "Stop"
#$VerbosePreference = "Continue"

# dot source the function files to load the functions into the session
. $PSScriptRoot\Shared.ps1

# check preconditions
if (-not ((Test-Path -Path $EpplusLibPath) -and ([System.IO.Path]::GetExtension($EpplusLibPath) -eq ".dll"))) {
} else {
    throw "Requires '${EpplusLibPath}'"
}

. $PSScriptRoot\Read-ExcelFile.ps1
. $PSScriptRoot\Get-ExcelCellContent.ps1
. $PSScriptRoot\Get-ExcelSpreadSheetTables.ps1
. $PSScriptRoot\Get-ExcelWorkSheets.ps1

$exportModuleMemberParams = @{
    Function = @(
        'Read-ExcelFile',
        'Get-ExcelCellContent',
        'Get-ExcelSpreadSheetTables',
        'Get-ExcelWorkSheets'
    )
    Variable = @(
        'EpplusLibPath'
    )
    Alias = "*"
}

Export-ModuleMember @exportModuleMemberParams
