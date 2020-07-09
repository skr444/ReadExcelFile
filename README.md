# Read-ExcelFile.psm1

This PowerShell script was created about the request to process an Excel file on a server where no Excel is installed.

The script is kept very simple. It is only about the import of an Excel table and the processing of the data into a PowerShell object which can be used for further processing.

PSGallery: [https://www.powershellgallery.com/packages/Read-ExcelFile/](https://www.powershellgallery.com/packages/Read-ExcelFile/)
## Installing the module
Start by installing the module
```powershell
install-module Read-ExcelFile
```

If you want to update the existing module run
```powershell
install-module Read-ExcelFile -Force
```
This will install the newest version from PSGallery

If you want to install a special version run
```powershell
install-module Read-ExcelFile -MinimumVersion 2.0
```
This will install minimum version 2.0

## Import Excel files
To import an Excel file simply run
```powershell
Read-ExcelFile -File [PATH_TO_YOUR_EXCEL_FILE]
```

If you want to import an Excel file with a specific worksheet run
```powershell
Read-ExcelFile -File [PATH_TO_YOUR_EXCEL_FILE] -WorkSheet "Table1"
```

## Working with the imported ExcelFile
Since the module converts the imported Excel file into a PowerShell object, you can easily continue working with the object.

To store the results in an variable run
```powershell
$items = Read-ExcelFile -File [PATH_TO_YOUR_EXCEL_FILE]
```
After that the variable can be reused e.g. 
```powershell
$items | where Status -eq "SOME_ITEM"
```

## Examples
You have a table like this:
| Status        | Id            |
| ------------- |---------------|
| new           | ID123445      |
| new           | ID123444      |
| closed        | ID123448      |

Your table is stored at: C:\Table\Table.xlsx

```powershell
install-module Read-ExcelFile -Force
$items = Read-ExcelFile -File C:\Table\Table.xlsx
$items | where Status -eq "closed
```
This will return

```powershell
Status: closed
Id: ID123448
```
