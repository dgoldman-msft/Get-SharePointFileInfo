function Get-TimeStamp {
    <#
        .SYNOPSIS
            Get a time stamp

        .DESCRIPTION
            Get a time date and time to create a custom time stamp

        .EXAMPLE
            None

        .NOTES
            Internal function
    #>

    [cmdletbinding()]
    param()
    return "[{0:MM/dd/yy} {0:HH:mm:ss}] -" -f (Get-Date)
}

function Save-Output {
    <#
    .SYNOPSIS
        Save output

    .DESCRIPTION
        Overload function for Write-Output

    .PARAMETER FailureObject
        Inbound failure log objects to be exported to csv

    .PARAMETER InputObject
        Inbound objects to be exported to csv

    .PARAMETER SaveFileOutput
        Flag for exporting the file object

    .PARAMETER SaveFailureOutput
        Flag for exporting the failure objects

    .PARAMETER SiteObject
        Inbound site objects list to be exported to csv

    .PARAMETER StringObject
        Inbound object to be printed and saved to log

    .EXAMPLE
        None

    .NOTES
        None
    #>

    [cmdletbinding()]
    param(
        [PSCustomObject]
        $FailureObject,

        [Object]
        $InputObject,

        [switch]
        $SaveFileOutput,

        [switch]
        $SaveFailureOutput,

        [PSCustomObject]
        $SiteObject,

        [Parameter(Mandatory = $True, Position = 0)]
        [string]
        $StringObject
    )

    process {
        try {
            Write-Output $StringObject
            if ($FailureObject -and $SaveFailureOutput.IsPresent) {
                $FailureObject | Export-Csv -Path (Join-Path -Path $LoggingDirectory -ChildPath $FailureLogSaveFileName) -Append -NoTypeInformation -ErrorAction Stop
                return
            }

            if ($SiteObject -and $SaveFileOutput.IsPresent) {
                $SiteObject | Export-Csv -Path (Join-Path -Path $LoggingDirectory -ChildPath $SharePointSiteListSaveFileName) -Append -NoTypeInformation -ErrorAction Stop
                return
            }

            if ($InputObject -and $SaveFileOutput.IsPresent) {
                $InputObject | Export-Csv -Path (Join-Path -Path $LoggingDirectory -ChildPath $FilesFoundLogName) -Append -NoTypeInformation -ErrorAction Stop
                return
            }

            # Console and log file output
            Out-File -FilePath (Join-Path -Path $LoggingDirectory -ChildPath $LoggingFileName) -InputObject $StringObject -Encoding utf8 -Append -ErrorAction Stop
        }
        catch {
            Save-Output "$(Get-TimeStamp) ERROR: $_"
            return
        }
    }
}
function Get-SharePointFileInfo {
    <#
        .SYNOPSIS
            Retrieve SharePoint file information

        .DESCRIPTION
            Connect to a SharePoint tenant and retrieve file information

        .PARAMETER EnableConsoleOutput
            Enable computer connection output to the console (noisy!)

        .PARAMETER FailureLogSaveFileName
            Failure log save file name

        .PARAMETER Filter
            Specifies the script block of the server-side filter to apply. See https://technet.microsoft.com/en-us/library/fp161380.aspx

        .PARAMETER IncludeOneDriveSites
            By default, the OneDrives are not returned. This switch includes all OneDrives.

        .PARAMETER LoggingDirectory
            Logging directory can be a local file share or network share with the necessary write permissions

        .PARAMETER LoggingFileName
            Script execution log file

        .PARAMETER RegisterPnPManagementAccess
            This will launch a device login flow that will ask you to consent to the application. Notice that is only required -once- per tenant. You will need to have appropriate access rights to be able to consent applications in your Azure AD.

        .PARAMETER SaveDataToDisk
            Switch to indicate you want to save results to a local or network location

        .PARAMETER SharePointSiteListSaveFileName
            SharePoint list including personal OneDrive locations to be exported to disk

        .PARAMETER TenantName
            Microsoft 365 Azure tenant name

        .EXAMPLE
            Get-SharePointFileInfo -TenantName "Contoso" -EnableConsoleOutput

            This will connect to a SharePoint site and grab the file list and the SharePoint sites and display them to the console

        .EXAMPLE
            Get-SharePointFileInfo -TenantName "Contoso" -EnableConsoleOutput -SaveToDisk

            This will connect to a SharePoint site and grab the file list from all SharePoint sites we have access to and save the log to disk

        .EXAMPLE
            Get-SharePointFileInfo -TenantName "Contoso" -EnableConsoleOutput -SaveToDisk -IncludeOneDriveSites

            This will connect to a SharePoint site and grab the tenant site list from all SharePoint sites and save the log to disk

        .NOTES
            For more information on PnP.PowerShell please see: https://pnp.github.io/powershell/articles/connecting.html
    #>

    [Diagnostics.CodeAnalysis.SuppressMessageAttribute("PSAvoidUsingConvertToSecureStringWithPlainText", "")]
    [OutputType('System.Object[]')]
    [CmdletBinding(DefaultParameterSetName = 'Default')]
    param (
        [switch]
        $EnableConsoleOutput,

        [string]
        $FailureLogSaveFileName = "FailuresLog.txt",

        [string]
        $FilesFoundLogName = 'SharePointFilesFound.csv',

        $Filter = "Url -like 'site'",

        [switch]
        $IncludeOneDriveSites,

        [string]
        $LoggingDirectory = 'C:\SharePointFileInfo',

        [string]
        $LoggingFileName = 'ScriptExecutionLogging.txt',

        [switch]
        $RegisterPnPManagementAccess,

        [switch]
        $SaveDataToDisk,

        [string]
        $SharePointSiteListSaveFileName = "SharePointSiteListWithOneDrive.csv",

        [string]
        $TenantName = "Default"
    )

    begin {
        $parameters = $PSBoundParameters
        $docLibrary = "Documents"
        [System.Collections.ArrayList]$failureEntries = @()
        [System.Collections.ArrayList]$fileList = @()
        if (-NOT( Test-Path -Path $LoggingDirectory )) {
            try {
                $null = New-Item -Path $LoggingDirectory -Type Directory -ErrorAction Stop
                Save-Output "$(Get-TimeStamp) Directory not found. Creating $($LoggingDirectory)"
            }
            catch {
                Save-Output "$(Get-TimeStamp) ERROR: $_"
                return
            }
        }

        Save-Output "$(Get-TimeStamp) Starting process"
    }

    process {
        try {
            if ($parameters.ContainsKey('RegisterPnPManagementAccess')) {
                Save-Output "$(Get-TimeStamp) Registering in your Azure tenant for Shell Access."
                Register-PnPManagementShellAccess
            }

            if ($TenantName -eq "Default") {
                Save-Output "$(Get-TimeStamp) ERROR: You have no specified a tenant name. Please enter a Tenant name and try again"
                return
            }

            Save-Output "$(Get-TimeStamp) Attempting to connect to: https://$TenantName-admin.sharepoint.com"
            if (-NOT ($connection = Connect-PnPOnline -Url "https://$TenantName-admin.sharepoint.com" -ReturnConnection)) {
                Save-Output "$(Get-TimeStamp) Unable to make a connection to $(https://$TenantName-admin.sharepoint.com)"
                return
            }
            else {
                Save-Output "$(Get-TimeStamp) Connection to successful"
                if ($parameters.ContainsKey('Verbose')) { $connection }
            }

            Save-Output "$(Get-TimeStamp) Obtaining SharePoint site list"

            if ($parameters.ContainsKey('IncludeOneDriveSites')) {
                Save-Output "$(Get-TimeStamp) Retrieving SharePoint site list for sites including OneDrive sites."
                $tenantSitesWithOneDrive = Get-PnPTenantSite -IncludeOneDriveSites -Filter "Url -like '-my.sharepoint.com/personal/'" -ErrorAction Stop -ErrorVariable Failed
            }

            Save-Output "$(Get-TimeStamp) Searching for SharePoint sites using the default filter: $($Filter)."
            $tenantSites = Get-PnPTenantSite -Filter $Filter -ErrorAction Stop -ErrorVariable Failed

            foreach ($site in $tenantSites) {
                if ($parameters.ContainsKey('EnableConsoleOutput')) { Save-Output "$(Get-TimeStamp) Attempting to connect to: $($site.url)" }

                Connect-PnPOnline -Url $site.Url -Credentials ($connection.PSCredential) -ErrorAction Stop -ErrorVariable Failed
                $filesFound = Get-PnPListItem -List $docLibrary -PageSize 1000 -ErrorAction SilentlyContinue -ErrorVariable Failed | Where-Object { $_["FileLeafRef"] -like "*.*" }

                ForEach ($file in $filesFound) {
                    if (-NOT ($file.FieldValues["SharedWithUsers"])) { $sharedWith = "Nobody" } else { $sharedWith = $file.FieldValues["SharedWithUsers"] }
                    if (-NOT ($file.FieldValues["CheckoutUser"])) { $checkoutUser = "Nobody" } else { $checkoutUser = $file.FieldValues["CheckoutUser"] }

                    $item = [PSCustomObject] @{
                        CheckoutUser        = $checkoutUser
                        CreatedByEmail      = $file.FieldValues["Author"].Email
                        CreatedTime         = $file.FieldValues["Created"]
                        DisplayName         = $file.FieldValues["_DisplayName"]
                        FileName            = $file.FieldValues["FileLeafRef"]
                        FileID              = $file.FieldValues["UniqueId"]
                        FileType            = $file.FieldValues["File_x0020_Type"]
                        FileSize_KB         = [Math]::Round(($file.FieldValues["File_x0020_Size"] / 1024), 2)
                        GUID                = $file.FieldValues["GUID"]
                        IsCurrentVersion    = $file.FieldValues["_IsCurrentVersion"]
                        IsCheckedoutToLocal = $file.FieldValues["IsCheckedoutToLocal"]
                        LastModifiedTime    = $file.FieldValues["Modified"]
                        ModifiedByEmail     = $file.FieldValues["Editor"].Email
                        SharedWithUsers     = $sharedWith
                        RelativeURL         = $file.FieldValues["FileRef"]
                    }

                    $null = $fileList.add($item)
                }
            }
        }
        catch {
            Save-Output "$(Get-TimeStamp) ERROR: $_"
            return
        }

        if ($parameters.ContainsKey('EnableConsoleOutput')) {
            if ($parameters.ContainsKey('IncludeOneDriveSites')) {
                Save-Output "$(Get-TimeStamp) SharePoint site list found!"
                $tenantSitesWithOneDrive | Select-Object Url
            }

            Save-Output "$(Get-TimeStamp) SharePoint files found!"
            $fileList | Format-Table
        }

        if ($Failed) {
            $failedEntry = [PSCustomObject]@{
                ComputerName = $Failed.OriginInfo.PSComputerName
                Time         = (Get-Date)
                Action       = $Failed.CategoryInfo.Activity
                Reason       = $Failed.CategoryInfo.Reason
            }
            $null = $failureEntries.add($failedEntry)
        }

        if ($failureEntries.count -gt 0) {
            Save-Output "$(Get-TimeStamp) WARNINGS / ERRORS: No logs found on some computers!" -FailureObjects $failureEntries -SaveFailureOutput:$True
            Save-Output "$(Get-TimeStamp) Please check $(Join-Path -Path $LoggingDirectory -ChildPath $LoggingFileName) for more information."
        }

        if ($parameters.ContainsKey('SaveDataToDisk')) {
            if ($parameters.ContainsKey('IncludeOneDriveSites')) {
                Save-Output "$(Get-TimeStamp) Exporting site list to $(Join-Path -Path $LoggingDirectory -ChildPath $SharePointSiteListSaveFileName). Please wait!" -SiteObject $tenantSitesWithOneDrive -SaveFileOutput:$True
            }
            Save-Output "$(Get-TimeStamp) Exporting registry logs to $(Join-Path -Path $LoggingDirectory -ChildPath $FilesFoundLogName). Please wait!" -InputObject $fileList -SaveFileOutput:$True
        }
    }

    end {
        Save-Output "$(Get-TimeStamp) Finished!"
    }
}