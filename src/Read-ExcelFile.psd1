@{
    # Assemblies that must be loaded prior to importing this module
    RequiredAssemblies = @('.\EPPlus.dll')

    # Script module or binary module file associated with this manifest.
    RootModule         = 'Read-ExcelFile.psm1'

    # Version number of this module.
    ModuleVersion      = '2.1'

    # ID used to uniquely identify this module
    GUID               = '5bc8dd3d-4ce9-45b6-b70c-d415bf36f4a1'

    # Author of this module
    Author             = 'Pascal Rimark'

    # Company or vendor of this module
    CompanyName        = 'Pascal Rimark'

    # Copyright statement for this module
    Copyright          = 'c 2020 All rights reserved.'

    # Description of the functionality provided by this module
    Description        = @'
PowerShell module to import Excel spreadsheets, without Excel.
Creates a PowerShell Object from an imported Excel spreadsheet.
'@



    # Functions to export from this module
    FunctionsToExport  = @(
        'Read-ExcelFile',
        'Get-ExcelCellContent',
        'Get-ExcelSpreadSheetTables',
        'Get-ExcelWorkSheets'
    )

    # Aliases to export from this module
    AliasesToExport    = @(
        'ref'
    )

    # Cmdlets to export from this module
    CmdletsToExport    = @()

    FileList           = @(
        '.\EPPlus.dll',
        '.\Read-ExcelFile.ps1'
        '.\Get-ExcelCellContent.ps1'
        '.\Get-ExcelSpreadSheetTables.ps1'
        '.\Get-ExcelWorkSheets.ps1'
    )

    # Private data to pass to the module specified in RootModule/ModuleToProcess
    PrivateData        = @{
        # PSData is module packaging and gallery metadata embedded in PrivateData
        # It's for rebuilding PowerShellGet (and PoshCode) NuGet-style packages
        # We had to do this because it's the only place we're allowed to extend the manifest
        # https://connect.microsoft.com/PowerShell/feedback/details/421837
        PSData = @{
            # The primary categorization of this module (from the TechNet Gallery tech tree).
            Category     = "Scripting Excel"

            # Keyword tags to help users find this module via navigations and search.
            Tags         = @("Excel", "EPPlus", "Export", "Import")

            # The web address of an icon which can be used in galleries to represent this module
            #IconUri = ""

            # The web address of this module's project or support homepage.
            ProjectUri   = ""

            # The web address of this module's license. Points to a page that's embeddable and linkable.
            #LicenseUri   = ""

            # Release notes for this particular version of the module
            #ReleaseNotes = $True

            # If true, the LicenseUrl points to an end-user license (not just a source license) which requires the user agreement before use.
            # RequireLicenseAcceptance = ""

            # Indicates this is a pre-release/testing version of the module.
            IsPrerelease = 'False'
        }
    }

    # Minimum version of the Windows PowerShell engine required by this module
    # PowerShellVersion = ''

    # Name of the Windows PowerShell host required by this module
    # PowerShellHostName = ''

    # Minimum version of the Windows PowerShell host required by this module
    # PowerShellHostVersion = ''

    # Minimum version of Microsoft .NET Framework required by this module
    # DotNetFrameworkVersion = ''

    # Minimum version of the common language runtime (CLR) required by this module
    # CLRVersion = ''

    # Processor architecture (None, X86, Amd64) required by this module
    # ProcessorArchitecture = ''

    # Modules that must be imported into the global environment prior to importing this module
    # RequiredModules = @()

    # Script files (.ps1) that are run in the caller's environment prior to importing this module.
    # ScriptsToProcess = @()

    # Type files (.ps1xml) to be loaded when importing this module
    # TypesToProcess = @()

    # Format files (.ps1xml) to be loaded when importing this module
    # FormatsToProcess = @()

    # Modules to import as nested modules of the module specified in RootModule/ModuleToProcess
    # NestedModules = @()

    # List of all modules packaged with this module
    # ModuleList = @()

    # List of all files packaged with this module
    # Variables to export from this module
    #VariablesToExport = '*'

    # HelpInfo URI of this module
    # HelpInfoURI = ''

    # Default prefix for commands exported from this module. Override the default prefix using Import-Module -Prefix.
    # DefaultCommandPrefix = ''

}