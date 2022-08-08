function Get-ExcelCellContent {
    <#
        .SYNOPSIS
        Returns the cell content from an Excel spreadsheet.
    
        .DESCRIPTION
        Returns the cell content from an Excel spreadsheet.

        .PARAMETER File
        Specify the file location of the excel file to import
        String

        .PARAMETER Address
        Specify the address of the cell (e.g A1 or B21). 
        String

        .EXAMPLE
        Get-ExcelCellContent .\MyExcel.xlsx -Address "B21"
    #>

    param(

        [Parameter(Mandatory = $true)]
        [string]$File,

        [Parameter(Mandatory = $true)]
        [string]$Address
    )

    Write-Verbose "ScriptRoot: $PSScriptRoot"

    try {
        $epplus = [System.Reflection.Assembly]::LoadFile("$PSScriptRoot\EPPlus.dll");
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

    try {
        if([string]::IsNullOrEmpty($WorkSheetName)) {
            $Worksheet = $xlspck.Workbook.Worksheets[1]
        } else {
            $Worksheet = $xlspck.Workbook.Worksheets["$WorkSheetName"]
        }
    } catch {
        $stream.Dispose()
        $xlspck.Stream.Close()
        $xlspck.Dispose()
        throw "FAILED_OPENING_WORKSHEET($WorkSheetName) - $($_.Exception.Message)"
    }

    return $Worksheet.Cells.Where({$_.Address -eq "$Address"}).Text
}
