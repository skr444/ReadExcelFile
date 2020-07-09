<#

    .SYNOPSIS
    Reads an Excel file and creates a PowerShell object from it.
 
    .DESCRIPTION
    Reads an Excel file and creates a PowerShell object from it.

    .NOTES
    File Name : Read-ExcelFile.ps1
    Author    : Pascal Rimark
    Requires  : PowerShell Version 3.0
    
    .LINK
    To provide feedback or for further assistance email:
    pascal@rimark.de

    .PARAMETER File
    Specify the file location of the excel file to import
    String

    .PARAMETER WorkSheetName
    Specify the name of the worksheet where the table to be imported is located. 
    String

    .EXAMPLE
    Read-ExcelFile .\MyExcel.xlsx
    .EXAMPLE
    Read-ExcelFile -File .\MyExcel.xlsx -WorkSheet "Table 2"
    .EXAMPLE
    Read-ExcelFile -File .\MyExcel.xlsx -WorkSheet "Table 2" -Verbose

#>

    param(
        [Parameter(Mandatory=$True)]
        [string]$File,
        [string]$WorkSheetName
    )

    $stopwatch =  [system.diagnostics.stopwatch]::StartNew()

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
        $stream.Dispose()
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

    Write-Verbose "WorkSheet is $($Worksheet.Name)"

    $Start = $Worksheet.Dimension.Start
    Write-Verbose "Dimension StartAddress: $($Start.Address)"

    $End = $Worksheet.Dimension.End   
    Write-Verbose "Dimension EndAddress: $($End.Address)"

    $headers = @()
    $export = @()

    for ($r = $Start.Row; $r -le $End.Row; $r++) {
        if($r -eq 1) {
            for ($c = $Start.Column; $c -le $End.Column; $c++) {
                $headers += $Worksheet.Cells[$r,$c]
                Write-Verbose "Header added $($Worksheet.Cells[$r,$c].Text)"
            }
        } else {
            $items = @()
            $rowItem = New-Object -TypeName psobject
            for ($c = $Start.Column; $c -le $End.Column; $c++) {
                $items += $Worksheet.Cells[$r,$c]
            }
            $index = 0
            foreach($h in $headers) {
                try {
                    if([Regex]::Matches($items[$index].Value,"(\d{2}.\d{2}.\d{4} \d{2}:\d{2}:\d{2})").Success) {
                        $t = [datetime]$items[$index].Value
                    } else {
                        $t = $items.Text[$index]
                    }
                    $rowItem | Add-Member -MemberType NoteProperty $h.Text $t -ErrorAction SilentlyContinue
                    Write-Verbose "RowItem added [Header:$($h.Text)] [Item:$($items.Text[$index])]"
                } catch {
                    Write-Verbose "Empty Row Detected [ROW:$r]"
                }
                $index++
            }
            $items = $null
            $export += $rowItem
            $rowItem = $null
        }
    }

    $stream.Dispose()
    $xlspck.Stream.Close()
    $xlspck.Dispose()

    Write-Verbose "Processed Items: $($export.Count)"
    $stopwatch.Stop()
    Write-Verbose "Elapsed Time: $($stopwatch.Elapsed)"

    return $export

#Read-ExcelFile -File C:\users\primark\Desktop\Powershell\impexc\O365-TeamsTelefonie.xlsx