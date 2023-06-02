# 1.0 setup a scheduled task

# 1.1 create variables

$jobName = "serso update pc"
# $dir=Get-ChildItem -Path 'C:\Program Files\SentinelOne\*\Sentinelctl.exe'

# 1.2 check if task already exists

if (Get-ScheduledTask -TaskName $jobName -ErrorAction SilentlyContinue) {
    Write-Host "Automatic updates task is already running in the background."
} 

# 1.3 if task doesn't exists

else {

    Write-Host "setting up scheduled task for serso master script"

    # Create a scheduled task that runs the script at user logon with highest privileges
    $action = New-ScheduledTaskAction -Execute "powershell.exe" -Argument "-ExecutionPolicy Bypass -NoExit -File C:\Temp\DJ_Serso_update_script.ps1"
    $trigger = New-ScheduledTaskTrigger -AtLogOn -User "$env:COMPUTERNAME\serso"
    $settings = New-ScheduledTaskSettingsSet -Priority 4 -AllowStartIfOnBatteries -DontStopIfGoingOnBatteries -WakeToRun
    Register-ScheduledTask -TaskName $jobName -Action $action -Trigger $trigger -Settings $settings -RunLevel Highest -Force -Description "rename machine"

    Write-Host "task has been scheduled on login of local serso"

    # Run the script for the first time
    & "C:\Temp\DJ_Serso_update_script.ps1"
}

Write-Host "starting windows updates"

# 2.0 update windows to latest version

# CD $dir.directory
# C:\Program files\sentinelone\*\Sentinelctl.exe unload -a

# 2.1 create variable

$UpdateSession = New-Object -ComObject Microsoft.Update.Session
$UpdateSearcher = $UpdateSession.CreateUpdateSearcher()
$SearchResult = $UpdateSearcher.Search("IsInstalled=0 and Type='Software'")

# 2.2 check for updates

if ($SearchResult.Updates.Count -eq 0) {
    Write-Host "No updates available."
}

# 2.3 install updates if there are some

else {
    Write-Host "Found $($SearchResult.Updates.Count) updates."
    $UpdatesToDownload = New-Object -ComObject Microsoft.Update.UpdateColl
    foreach ($Update in $SearchResult.Updates) {
        Write-Host "$update added"
        $UpdatesToDownload.Add($Update)
    }
    $Downloader = $UpdateSession.CreateUpdateDownloader()
    $Downloader.Updates = $UpdatesToDownload
    $Downloader.Download()

    Write-Host "installing updates"

    $Installer = New-Object -ComObject Microsoft.Update.Installer
    $Installer.Updates = $UpdatesToDownload
    $Result = $Installer.Install()
    if ($Result.ResultCode -eq 2) {
        Write-Host "Updates installed successfully."
        Restart-Computer -Force
    } else {
        Write-Host "Installation failed with result code $($Result.ResultCode)."
    }
}

# start optional updates
# create log path
New-Item -ItemType directory -Path C:\Drivers -ErrorAction SilentlyContinue

# Install prerequisite
Install-PackageProvider -Name NuGet -MinimumVersion 2.8.5.021 -Force
Install-Module -Name PSWindowsUpdate -Force
Import-Module -Name PSWindowsUpdate -Force

# Register to MS Update Service
Add-WUServiceManager -ServiceID "7971f918-a847-4430-9279-4a52d1efe18d" -Confirm:$false

# Download and install drivers and repeat once.
Install-WindowsUpdate -Install -AcceptAll -UpdateType Driver -MicrosoftUpdate -ForceDownload -ForceInstall -IgnoreReboot -ErrorAction SilentlyContinue | Out-File "c:\Drivers\Drivers_Install_1_$(get-date -f dd-MM-yyyy).log" -Force
Install-WindowsUpdate -Install -AcceptAll -UpdateType Driver -MicrosoftUpdate -ForceDownload -ForceInstall -IgnoreReboot -ErrorAction SilentlyContinue | Out-File "c:\Drivers\Drivers_Install_2_$(get-date -f dd-MM-yyyy).log" -Force

Write-Host "windows updates have been completed and is now up to date"
Write-Host "starting the removal process of microsoft app bloatware"

# CD $dir.directory
# C:\Program files\sentinelone\*\Sentinelctl.exe load -a

# 3.0 remove pre-installed microsoft store apps

Set-ExecutionPolicy Unrestricted -Scope CurrentUser -Force

get-appxpackage *instagram* | remove-appxpackage
get-appxpackage *spotify* | remove-appxpackage
get-appxpackage *messenger* | remove-appxpackage
get-appxpackage *solitairecollection* | remove-appxpackage
get-appxpackage *tiktok* | remove-appxpackage
get-appxpackage *disney* | remove-appxpackage
get-appxpackage *primevideo* | remove-appxpackage
get-appxpackage *netflix* | remove-appxpackage
get-appxpackage *zunemusic* | remove-appxpackage
get-appxpackage *maps* | remove-appxpackage
get-appxpackage *skypeapp* | remove-appxpackage
get-appxpackage *friends* | remove-appxpackage
get-appxpackage *candycrush* | remove-appxpackage
get-appxpackage *spades* | remove-appxpackage
get-appxpackage *mahjong* | remove-appxpackage

# 4.0 download software the costumer wants

Set-ExecutionPolicy Bypass -Scope Process -Force; [System.Net.ServicePointManager]::SecurityProtocol = [System.Net.ServicePointManager]::SecurityProtocol -bor 3072; iex ((New-Object System.Net.WebClient).DownloadString('https://community.chocolatey.org/install.ps1'))

Start-Sleep -Seconds 120
choco feature enable -n allowGlobalConfirmation
$wshell = New-Object -ComObject Wscript.Shell

$answer = $wshell.Popup("Do you want to install Adobe Reader?",0,"Confirmation",0x1)
if($answer -eq 1){
    choco install adobereader -y --yes --confirm
}

$answer = $wshell.Popup("Do you want to install Java Runtime?",0,"Confirmation",0x1)
if($answer -eq 1){
    choco install javaruntime -y --yes --confirm
}

$answer = $wshell.Popup("Do you want to install 7-Zip?",0,"Confirmation",0x1)
if($answer -eq 1){
    choco install 7zip -y --yes --confirm
}

$answer = $wshell.Popup("Do you want to install Google Chrome?",0,"Confirmation",0x1)
if($answer -eq 1){
    choco install googlechrome -y --yes --confirm
}

$answer = $wshell.Popup("Do you want to install TeamViewer?",0,"Confirmation",0x1)
if($answer -eq 1){
    choco install teamviewer -y --yes --confirm
}

choco uninstall chocolatey -y

# 5.0 unschedule the task 


Import-Module ScheduledTasks
Unregister-scheduledTask *serso* -Confirm:$false
Write-Host "scheduled task has been removed"

# 6.0 check if bitlocker has encrypted the c drive

# 6.1 create variables

$teller = 1

# start check

while($teller -eq 1){
    $running = manage-bde -status C: | Select-Object -Index 7
    $newrunning = $running.substring(26,1)
    $data = manage-bde -status C: | Select-Object -Index 9
    $newdata = $data.substring(26,3)
    if($newrunning -eq "N"){
        Write-Host "bitlocker not installed"
        $teller = 0
    }
    elseif($newdata -eq "100"){
        Write-Host "bitlocker encryption is ready you can execute a bios upgrade"
        $teller = 0
    }
    elseif($newdata -ne "100"){
        Write-Host "encryption not ready yet encryption at $newdata"
        start-sleep -Seconds 100
    }
}

# 7.0 rename host and final restart

Add-Type -AssemblyName Microsoft.VisualBasic
$name = [Microsoft.VisualBasic.Interaction]::InputBox('if already changed for example with windows 11 first boot enter empty field.', 'device name', "")

start-sleep -seconds 5

if($name -eq ""){
    Write-Host "check done"
    Remove-Item -Path "C:\Temp\DJ_Serso_update_script.ps1"
    Write-Host "device is ready."
    Restart-Computer -Force
}

Write-Host "computer name changes to $name"
Write-Host "final restart in"

start-sleep -Seconds 5

Remove-Item -Path "C:\Temp\DJ_Serso_update_script.ps1"
Rename-Computer -NewName $name -Force -Restart
            