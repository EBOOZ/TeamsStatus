<#
.NOTES
Name: Get-TeamsStatus.ps1
Author: Danny de Vries
Requires: PowerShell v2 or higher
Version History: https://github.com/EBOOZ/TeamsStatus/commits/main
.SYNOPSIS
Sets the status of the Microsoft Teams client to Home Assistant.
.DESCRIPTION
This script is monitoring the Teams client logfile for certain changes. It
makes use of two sensors that are created in Home Assistant up front.
sensor.teams_status displays that availability status of your Teams client based 
on the icon overlay in the taskbar on Windows. sensor.teams_activity shows if you 
are in a call or not based on the App updates deamon, which is paused as soon as 
you join a call.
#>
# Configure the varaibles below that will be used in the script
$HAToken = "<Insert token>" # Example: eyJ0eXAiOiJKV1...
$UserName = "<UserName>" # When not sure, open a command prompt and type: echo %USERNAME%
$HAUrl = "<HAUrl>" # Example: https://yourha.duckdns.org

# Don't edit the code below, unless you want to change the value language
$headers = @{"Authorization"="Bearer $HAToken";}
$Enable = 1
$CurrentStatus = "Offline"
DO {
# Get Teams Logfile and last icon overlay status
$TeamsStatus = Get-Content -Path "C:\Users\$UserName\AppData\Roaming\Microsoft\Teams\logs.txt" -Tail 100 | Select-String -Pattern `
  'Setting the taskbar overlay icon -',`
  'StatusIndicatorStateService: Added',`
  'Main window is closing',`
  'main window closed' | Select-Object -Last 1
# Get Teams Logfile and last app update deamon status
$TeamsActivity = Get-Content -Path "C:\Users\$UserName\AppData\Roaming\Microsoft\Teams\logs.txt" -Tail 100 | Select-String -Pattern `
  'Resuming daemon App updates',`
  'Pausing daemon App updates',`
  'SfB:TeamsNoCall',`
  'SfB:TeamsPendingCall',`
  'SfB:TeamsActiveCall' | Select-Object -Last 1

If ($TeamsStatus -like "*Setting the taskbar overlay icon - Available*" -or `
    $TeamsStatus -like "*StatusIndicatorStateService: Added Available*") {
    $Status = "Available"
    Write-Host $Status
}
ElseIf ($TeamsStatus -like "*Setting the taskbar overlay icon - Busy*" -or `
        $TeamsStatus -like "*StatusIndicatorStateService: Added Busy*" -or `
        $TeamsStatus -like "*Setting the taskbar overlay icon - On the phone*") {
    $Status = "Busy"
    Write-Host $Status
}
ElseIf ($TeamsStatus -like "*Setting the taskbar overlay icon - Away*" -or `
        $TeamsStatus -like "*StatusIndicatorStateService: Added Away*") {
    $Status = "Away"
    Write-Host $Status
}
ElseIf ($TeamsStatus -like "*Setting the taskbar overlay icon - Do not disturb *" -or `
        $TeamsStatus -like "*StatusIndicatorStateService: Added DoNotDisturb*" -or `
        $TeamsStatus -like "*Setting the taskbar overlay icon - Focusing*" -or `
        $TeamsStatus -like "*StatusIndicatorStateService: Added Focusing*") {
    $Status = "Do not disturb"
    Write-Host $Status
}
ElseIf ($TeamsStatus -like "*Setting the taskbar overlay icon - In a meeting*" -or `
        $TeamsStatus -like "*StatusIndicatorStateService: Added InAMeeting*") {
    $Status = "In a meeting"
    Write-Host $Status
}
ElseIf ($TeamsStatus -like "*ain window*") {
    $Status = "Offline"
    Write-Host $Status
}

If ($TeamsActivity -like "*Resuming daemon App updates*" -or `
    $TeamsActivity -like "*SfB:TeamsNoCall*") {
    $Activity = "Not in a call"
    $ActivityIcon = "mdi:phone-off"
    Write-Host $Activity
}
ElseIf ($TeamsActivity -like "*Pausing daemon App updates*" -or `
        $TeamsActivity -like "*SfB:TeamsActiveCall*") {
    $Activity = "In a call"
    $ActivityIcon = "mdi:phone-in-talk-outline"
    Write-Host $Activity
}

If ($CurrentStatus -ne $Status) {
    $CurrentStatus = $Status

    $params = @{
     "state"="$CurrentStatus";
     "attributes"= @{
        "friendly_name"="Microsoft Teams status";
        "icon"="mdi:microsoft-teams";
        }
     }
    Invoke-RestMethod -Uri "$HAUrl/api/states/sensor.teams_status" -Method POST -Headers $headers -Body ($params|ConvertTo-Json) -ContentType "application/json" 
}

If ($CurrentActivity -ne $Activity) {
    $CurrentActivity = $Activity

    $params = @{
     "state"="$Activity";
     "attributes"= @{
        "friendly_name"="Microsoft Teams activity";
        "icon"="$ActivityIcon";
        }
     }
    Invoke-RestMethod -Uri "$HAUrl/api/states/sensor.teams_activity" -Method POST -Headers $headers -Body ($params|ConvertTo-Json) -ContentType "application/json" 
}
    Start-Sleep 1
} Until ($Enable -eq 0)
