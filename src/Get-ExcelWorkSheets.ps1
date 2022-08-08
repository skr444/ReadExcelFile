<#

    .SYNOPSIS
    Returns the cell content from an Excel spreadsheet.
 
    .DESCRIPTION
    Returns the cell content from an Excel spreadsheet.

    .NOTES
    File Name : Get-ExcelCellContent.ps1
    Author    : Pascal Rimark
    Requires  : PowerShell Version 3.0
    
    .LINK
    To provide feedback or for further assistance email:
    pascal@rimark.de

    .PARAMETER File
    Specify the file location of the excel file to import
    String

    .PARAMETER Address
    Specify the address of the cell (e.g A1 or B21). 
    String

    .EXAMPLE
    Get-ExcelCellContent .\MyExcel.xlsx -Address "B21"

#>

function Get-ExcelWorkSheets() {
    param(

        [Parameter(Mandatory=$True)]
        [string]$File = "C:\users\primark\desktop\UICS-O365-Pre-RoleoutV2.xlsx"

    )

    $stopwatch =  [system.diagnostics.stopwatch]::StartNew()

    Write-Verbose "ScriptRoot: $PSScriptRoot"

    try {
        $epplus = [System.Reflection.Assembly]::LoadFile("$PSScriptRoot\EPPlus.dll");
        #$epplus = [System.Reflection.Assembly]::LoadFile("C:\Users\primark\Desktop\Powershell\impexc\EPPlus.dll");
        Write-Verbose "Assembly loaded"
    } catch {
        throw "FAILED_LOADING_ASSEMBLY_FILE - $($_.Exception.Message)"
    }

    try {
        $stream = New-Object -TypeName System.IO.FileStream -ArgumentList $File, 'Open', 'Read', 'ReadWrite'
        Write-Verbose "Stream opened ($($stream.Name))"
    } catch {
        throw "FAILED_CREATING_FILESTREAM - $($_.Exception.Message)"
    }

    try {
        $xlspck = New-Object OfficeOpenXml.ExcelPackage
        $xlspck.Load($stream)
        Write-Verbose "Package loaded - Loaded from stream"
    } catch {
        $stream.Dispose()
        $xlspck.Stream.Close()
        $xlspck.Dispose()
        throw "FAILED_CREATING EXCELPACKAGE - $($_.Exception.Message)"
    }

    return $xlspck.Workbook.Worksheets.Name
}