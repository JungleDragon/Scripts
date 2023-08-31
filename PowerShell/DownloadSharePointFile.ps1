<#Create New Registered App in SharePoint at https://tenant.sharepoint.com/sites/marketing/_layouts/15/AppRegNew.aspx
Make sure to copy Client ID and Client Secret

Title: SharePoint Read for PnP.PowerShell Module
App Domain: localhost
Redirect URL: https://localhost
Click "Create"

Go to https://tenant.sharepoint.com/sites/marketing/_layouts/15/AppInv.aspx
Lookup App ID using Client ID
Permission Request XML:
<AppPermissionRequests AllowAppOnlyPolicy="true">  
   <AppPermissionRequest Scope="http://sharepoint/content/sitecollection/web/list" 
    Right="FullControl" />
</AppPermissionRequests>

If you need to delete the App or re-register, you can go to https://tenant.sharepoint.com/sites/marketing/_layouts/15/AppPrincipals.aspx?Scope=Web
#>

Set-ExecutionPolicy RemoteSigned -Force

Write-Host "Checking if ExchangeOnlineManagement Module is already installed"
try
	{Write-Host "Importing PnP.PowerShell Module"
	Import-Module PnP.PowerShell -ErrorAction Stop
	Write-Host "PnP.PowerShell Module Imported"}
catch
	{Write-Host "Need to install PnP.PowerShell Module"
	Install-Module -Name PnP.PowerShell -Force
	Write-Host "PnP.PowerShell Module Installed"
	
	Write-Host "Importing PnP.PowerShell Module"
	Import-Module PnP.PowerShell
	Write-Host "PnP.PowerShell Module Imported"}

# The URL can be something like https://example.sharepoint.com/sites/BI
$SiteURL = "https://example.sharepoint.com/sites/BI"

# Connect to the PNP module using the variables previously informed
Connect-PnPOnline -Url $SiteURL -ClientId "00000000-0000-0000-0000-0000000" -ClientSecret "0000000000000000000000" -WarningAction Ignore

# Defines the directory where the file will be downloaded
$DownloadPath = "C:\Users\User\Downloads"

# Path to file from SP folder
$FileRelativeUrl = "/Documents/Documents/File.ext"

# Name of the file that is going to be downloaded and renamed
$FileName = "File.ext"

# Download the file from the Sharepoint
Get-PnPFile -Url $FileRelativeUrl -Path $DownloadPath -FileName $FileName -AsFile -Force

# Disconnects from PnP module
Disconnect-PnPOnline
