#Search for File
Get-Childitem -Path C:\ -Include "*file name*" -Recurse -ErrorAction SilentlyContinue -Force

#Search for a String in a File
Get-Childitem -Path C:\ -Include "*file name*" -Recurse -ErrorAction SilentlyContinue -Force |
Select-String -Pattern 'string'

#Delete specific files
Get-Childitem -Path C:\ -Include "*file name*" -Recurse -ErrorAction SilentlyContinue | Remove-Item

#Check Folder Size
Get-ChildItem -Path "C:\Temp" -Recurse -ErrorAction SilentlyContinue | 
    Measure-Object -Property Length -Sum | 
    Select-Object Sum, Count

#Compress Folder
Compress-Archive -Path C:\Reference -DestinationPath C:\Archives\Draft.zip

#Expand Folder
Expand-Archive -Path Draft.zip

#Install MSI quietly and wait for it to finish before running next command
Start-Process msiexec.exe -Wait -ArgumentList '/I C:\Windows\Temp\SQLIO.msi /quiet'

#Check for ADS (Run from the root of a drive you want to check)
gci -recurse | % { gi $_.FullName -stream * } | where stream -ne ':$Data'

#List Programs Installed on Machine
Get-WmiObject -Class Win32_Product | Select-Object -Property Name

#Uninstall Program
$MyApp = Get-WmiObject -Class Win32_Product | Where-Object{$_.Name -eq "Name"}
$MyApp.Uninstall()

#Get Wi-Fi Password
Netsh wlan show profile name=”SSID Name” key=clear
