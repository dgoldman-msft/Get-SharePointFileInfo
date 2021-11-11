# Get-SharePointFileInfo

Script for connecting to a SharePoint site and getting a file list of all files found including site list

> EXAMPLE 1: Get-SharePointFileInfo -TenantName "Contoso" -EnableConsoleOutput

- This will connect to a SharePoint site and grab the file list and the SharePoint sites and display them to the console

> EXAMPLE 2: Get-SharePointFileInfo -TenantName "Contoso" -EnableConsoleOutput -SaveToDisk

- This will connect to a SharePoint site and grab the file list from all SharePoint sites we have access to and save the log to disk

> EXAMPLE 3: Get-SharePointFileInfo -TenantName "Contoso" -EnableConsoleOutput -SaveToDisk -IncludeOneDriveSites

- This will connect to a SharePoint site and grab the tenant site list from all SharePoint sites and save the log to disk

NOTE: All logs will be saved to the following variable -> $EventLogSaveLocation = 'c:\SharePointFileInfo'. This can be a local share or a network share with the correct write permissions
