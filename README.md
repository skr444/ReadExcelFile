# ReadExcelFile.psm1

This PowerShell script was created about the request to process an Excel file on a server where no Excel is installed.

The script is kept very simple. It is only about the import of an Excel table and the processing of the data into a PowerShell object which can be used for further processing.

PSGallery: not available

## PowerShell version

This module requires at least PowerShell 5.1

## Installing the module
### From PowerShell Gallery
Start by installing the module
```powershell
Install-Module ReadExcelFile
```

If you want to update the existing module run
```powershell
Install-Module ReadExcelFile -Force
```
This will install the latest version from PSGallery

If you want to install a special version run
```powershell
Install-Module ReadExcelFile -MinimumVersion 2.0
```
This will install minimum version 2.0

### From Git repository
If you want to build and deploy the module manually after cloning the git repository, run
```powershell
.\build.ps1
```

This will create the module in .\build\ReadExcelFile

To find out in which PowerShell profile folder to copy the built module, run
```powershell
Get-Module -ListAvailable
```

This will list all installed modules available for import along with the profile.

Copy the entire module folder into the Modules folder in the desired PowerShell profile, which is usually the user profile returned by
```powershell
$PROFILE.CurrentUserCurrentHost
```

This returns the PowerShell profile script in which the Modules folder is located.

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
Install-Module ReadExcelFile -Force
$items = Read-ExcelFile -File C:\Table\Table.xlsx
$items | where Status -eq "closed
```
This will return

```powershell
Status: closed
Id: ID123448
```
