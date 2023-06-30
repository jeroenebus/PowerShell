<#
.Synopsis
    Script to install a different system language for the user.
    Installing language packs requires administrive permissions.
    This script will make use of a scheduled task to change the language for the current user.
.DESCRIPTION
    This script will change the system language of the machine for the current/ new user.
.EXAMPLE
    You can customize the script by changing the organization name, language and region.
.NOTES
    Filename: Install-LanguageExperiencePack_Windows11fr-CA.ps1
    Author: Jeroen Ebus (https://manage-the.cloud)
    	    Part of the script by: Peter Klapwijk - www.inthecloud247.com
	    Part of the script by: Olivier Kieselbach - www.oliverkieselbach.com
    Modified date: 2023-06-29
    Version 1.0 - Release notes/details
#>

# Microsoft Intune Management Extension might start a 32-bit PowerShell instance. If so, restart as 64-bit PowerShell.
If ($ENV:PROCESSOR_ARCHITEW6432 -eq "AMD64") {
    Try {
        &"$ENV:WINDIR\SysNative\WindowsPowershell\v1.0\PowerShell.exe" -File $PSCOMMANDPATH
    }
    Catch {
        Throw "Failed to start $PSCOMMANDPATH"
    }
    Exit
}

# Set variables:
# Company name.
$CompanyName = "Manage The Cloud"
# The language we want as new default. Language tag can be found here: https://learn.microsoft.com/en-us/windows-hardware/manufacture/desktop/available-language-packs-for-windows
$language = "fr-CA"
# Geographical ID we want to set. GeoID can be found here: https://learn.microsoft.com/en-us/windows/win32/intl/table-of-geographical-locations?redirectedfrom=MSDN
$geoId = "39"  # Canada.

# Start Transcript.
Start-Transcript -Path "$env:ProgramData\$CompanyName\Logs\$($(Split-Path $PSCommandPath -Leaf).ToLower().Replace(".ps1",".log"))" | Out-Null

# Custom folder for temp scripts.
"... Creating custom temp script folder"
$scriptFolderPath = "$env:SystemDrive\ProgramData\$CompanyName\CustomTempScripts"
New-Item -ItemType Directory -Force -Path $scriptFolderPath
"`n"

$userConfigScriptPath = $(Join-Path -Path $scriptFolderPath -ChildPath "UserConfig.ps1")
"... Creating userconfig scripts"
# We could encode the complete script to prevent the escaping of $, but I found it easier to maintain.
# To not encode. I do not have to decode/encode all the time for modifications.
$userConfigScript = @"
`$language = "$language"
Start-Transcript -Path `$env:TEMP"\LXP-UserSession-Config-$language.log" | Out-Null
`$geoId = "$geoId"
# Important for regional change like date and time...
"`Set-WinUILanguageOverride = $language`"
Set-WinUILanguageOverride -Language `"$language"
Set-WinUserLanguageList -LanguageList `"$language" -Force
`$OldList` = Get-WinUserLanguageList
`$UserLanguageList = New-WinUserLanguageList -Language `"$language"
`$UserLanguageList += `$OldList | Where-Object { `$_.LanguageTag -ne `"$language" }
"Setting new user language list:"
`$UserLanguageList | Select-Object LanguageTag
"Set-WinUserLanguageList -LanguageList ..."
Set-WinUserLanguageList -LanguageList `$UserLanguageList -Force
"`Set-Culture = $language`"
Set-Culture -CultureInfo `"$language"
"`Set-WinHomeLocation = $geoId`"
Set-WinHomeLocation -GeoId `"$geoId"
Stop-Transcript -Verbose
"@

$userConfigScriptHiddenStarterPath = $(Join-Path -Path $scriptFolderPath -ChildPath "UserConfigHiddenStarter.vbs")
$userConfigScriptHiddenStarter = @"
sCmd = "powershell.exe -ex bypass -file ""$userConfigScriptPath"""
Set oShell = CreateObject("WScript.Shell")
oShell.Run sCmd,0,true
"@

# Install an additional language pack including FODs.
"Installing languagepack"
Install-Language $language -CopyToSettings

# Set System Preferred UI Language.
"Set SystemPreferredUILanguage"
$language

# Check status of the installed language pack.
"Checking installed languagepack status"
$installedLanguage = (Get-InstalledLanguage).LanguageId

if ($installedLanguage -like $language) {
    Write-Host "Language $language installed"
}
else {
    Write-Host "Failure! Language $language NOT installed"
    Exit 1
}

# Check status of the System Preferred Language.
$SystemPreferredUILanguage = Get-SystemPreferredUILanguage

if ($SystemPreferredUILanguage -like $language) {
    Write-Host "System Preferred UI Language set to $language. OK"
}
else {
    Write-Host "Failure! System Preferred UI Language NOT set to $language. System Preferred UI Language is $SystemPreferredUILanguage"
    Exit 1
}

# Configure new language defaults under current user (system account) after which it can be copied to the system.
# Set Win UI Language Override for regional changes.
"Set WinUILanguageOverride"
Set-WinUILanguageOverride -Language $language

# Set Win User Language List, sets the current user language settings.
"Set WinUserLanguageList"
$OldList = Get-WinUserLanguageList
$UserLanguageList = New-WinUserLanguageList -Language $language
$UserLanguageList += $OldList | Where-Object { $_.LanguageTag -ne $language }
$UserLanguageList | Select-Object LanguageTag
Set-WinUserLanguageList -LanguageList $UserLanguageList -Force

# Set Culture, sets the user culture for the current user account.
"Set culture"
Set-Culture -CultureInfo $language

# Set Win Home Location, sets the home location setting for the current user.
"Set WinHomeLocation"
Set-WinHomeLocation -GeoId $geoId

# Copy User Internaltional Settings from current user to System, including Welcome screen and new user.
"Copy UserInternationalSettingsToSystem"
Copy-UserInternationalSettingsToSystem -WelcomeScreen $True -NewUser $True

# We have to switch the language for the current user session. The powershell cmdlets must be run in the current logged on user context.
# Creating a temp scheduled task to run on-demand in the current user context does the trick here.
"Trigger language change for current user session via ScheduledTask = LXP-UserSession-Config-$language"
Out-File -FilePath $userConfigScriptPath -InputObject $userConfigScript -Encoding ascii
Out-File -FilePath $userConfigScriptHiddenStarterPath -InputObject $userConfigScriptHiddenStarter -Encoding ascii

# REMARK: Usage of wscript as hidden starter may be blocked because of security restrictions like AppLocker, ASR, etc...
#         Switch to PowerShell if this represents a problem in your environment.
$taskName = "LXP-UserSession-Config-$language"
$action = New-ScheduledTaskAction -Execute "wscript.exe" -Argument """$userConfigScriptHiddenStarterPath"""
$trigger = New-ScheduledTaskTrigger -AtLogOn

# Grab the owner of explorer.
$proc = Get-CimInstance Win32_Process -Filter "name = 'explorer.exe'"
# You might have multiple instances of explorer running, so we pick the first one.
$domain = Invoke-CimMethod -InputObject $proc[0] -MethodName GetOwner | select-object -ExpandProperty Domain
$username = Invoke-CimMethod -InputObject $proc[0] -MethodName GetOwner | select-object -ExpandProperty User
# Combine the domain prefix with the username.
$fullupn = "$domain\$username"
$principal = New-ScheduledTaskPrincipal -UserId $fullupn
$settings = New-ScheduledTaskSettingsSet -AllowStartIfOnBatteries
$task = New-ScheduledTask -Action $action -Trigger $trigger -Principal $principal -Settings $settings
Register-ScheduledTask $taskName -InputObject $task
Start-ScheduledTask -TaskName $taskName 

Start-Sleep -Seconds 30

Unregister-ScheduledTask -TaskName $taskName -Confirm:$false

# Trigger 'LanguageComponentsInstaller\ReconcileLanguageResources' otherwise 'Windows Settings' need a long time to change finally.
"Trigger ScheduledTask = LanguageComponentsInstaller\ReconcileLanguageResources"
Start-ScheduledTask -TaskName "\Microsoft\Windows\LanguageComponentsInstaller\ReconcileLanguageResources"

Start-Sleep 10

# Trigger store updates, there might be new app versions due to the language change.
"Trigger MS Store updates for app updates"
Get-CimInstance -Namespace "root\cimv2\mdm\dmmap" -ClassName "MDM_EnterpriseModernAppManagement_AppManagement01" | Invoke-CimMethod -MethodName "UpdateScanMethod"

# Add registry key for MECM detection.
"Add registry key for MECM detection"
REG add "HKLM\Software\$CompanyName\LanguageXPWIN11\v1.0" /v "SetLanguage-$language" /t REG_DWORD /d 1

Exit 3010
Stop-Transcript
