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
.PARAMETER SetStatus
    Run the script with the SetStatus-parameter to set the status of Microsoft Teams
    directly from the commandline.
.EXAMPLE
    .\Get-TeamsStatus.ps1 -SetStatus "Offline"
#>
# Configuring parameter for interactive run
Param($SetStatus)

# Configure the variables below that will be used in the script
$HAToken = "<Insert token>" # Example: eyJ0eXAiOiJKV1...
$UserName = "<UserName>" # When not sure, open a command prompt and type: echo %USERNAME%
$HAUrl = "<HAUrl>" # Example: https://yourha.duckdns.org

# Set language variables below
$lgAvailable = "Available"
$lgBusy = "Busy"
$lgAway = "Away"
$lgBeRightBack = "Be right back"
$lgDoNotDisturb = "Do not disturb"
$lgInAMeeting = "In a meeting"
$lgOffline = "Offline"
$lgNotInACall = "Not in a call"
$lgInACall = "In a call"

# Set icons to use for call activity
$iconInACall = "mdi:phone-in-talk-outline"
$iconNotInACall = "mdi:phone-off"

################################################################
# Don't edit the code below, unless you know what you're doing #
################################################################
$headers = @{"Authorization"="Bearer $HAToken";}
$Enable = 1

# Run the script when a parameter is used and stop when done
If($null -ne $SetStatus){
    Write-Host ("Setting Microsoft Teams status to "+$SetStatus+":")
    $params = @{
     "state"="$SetStatus";
     "attributes"= @{
        "friendly_name"="Microsoft Teams status";
        "icon"="mdi:microsoft-teams";
        }
     }
	 
    $params = $params | ConvertTo-Json
    Invoke-RestMethod -Uri "$HAUrl/api/states/sensor.teams_status" -Method POST -Headers $headers -Body ([System.Text.Encoding]::UTF8.GetBytes($params)) -ContentType "application/json"
	
    break
}

# Clear the status and activity variable at the start
$CurrentStatus = ""
$CurrentActivity = ""

# Start monitoring the Teams logfile when no parameter is used to run the script
DO {
# Get Teams Logfile and last icon overlay status
$TeamsStatus = Get-Content -Path "C:\Users\$UserName\AppData\Roaming\Microsoft\Teams\logs.txt" -Tail 100 | Select-String -Pattern `
  'Setting the taskbar overlay icon -',`
  'StatusIndicatorStateService: Added' | Select-Object -Last 1
# Get Teams Logfile and last app update deamon status
$TeamsActivity = Get-Content -Path "C:\Users\$UserName\AppData\Roaming\Microsoft\Teams\logs.txt" -Tail 100 | Select-String -Pattern `
  'Resuming daemon App updates',`
  'Pausing daemon App updates',`
  'SfB:TeamsNoCall',`
  'SfB:TeamsPendingCall',`
  'SfB:TeamsActiveCall' | Select-Object -Last 1
# Get Teams application process
$TeamsProcess = Get-Process -Name Teams -ErrorAction SilentlyContinue

# Check if Teams is running and start monitoring the log if it is
If ($null -ne $TeamsProcess) {
    If ($TeamsStatus -like "*Setting the taskbar overlay icon - Available*" -or `
        $TeamsStatus -like "*StatusIndicatorStateService: Added Available*") {
        $Status = $lgAvailable
        Write-Host $Status
    }
    ElseIf ($TeamsStatus -like "*Setting the taskbar overlay icon - Busy*" -or `
            $TeamsStatus -like "*StatusIndicatorStateService: Added Busy*" -or `
            $TeamsStatus -like "*Setting the taskbar overlay icon - On the phone*" -or `
            $TeamsStatus -like "*StatusIndicatorStateService: Added OnThePhone*") {
        $Status = $lgBusy
        Write-Host $Status
    }
    ElseIf ($TeamsStatus -like "*Setting the taskbar overlay icon - Away*" -or `
            $TeamsStatus -like "*StatusIndicatorStateService: Added Away*") {
        $Status = $lgAway
        Write-Host $Status
    }
    ElseIf ($TeamsStatus -like "*Setting the taskbar overlay icon - $lgBeRightBack*" -or `
            $TeamsStatus -like "*StatusIndicatorStateService: Added BeRightBack*") {
        $Status = $lgBeRightBack
        Write-Host $Status
    }
    ElseIf ($TeamsStatus -like "*Setting the taskbar overlay icon - Do not disturb *" -or `
            $TeamsStatus -like "*StatusIndicatorStateService: Added DoNotDisturb*" -or `
            $TeamsStatus -like "*Setting the taskbar overlay icon - Focusing*" -or `
            $TeamsStatus -like "*StatusIndicatorStateService: Added Focusing*") {
        $Status = $lgDoNotDisturb
        Write-Host $Status
    }
    ElseIf ($TeamsStatus -like "*Setting the taskbar overlay icon - In a meeting*" -or `
            $TeamsStatus -like "*StatusIndicatorStateService: Added InAMeeting*") {
        $Status = $lgInAMeeting
        Write-Host $Status
    }
    ElseIf ($TeamsStatus -like "*Setting the taskbar overlay icon - $lgOffline*" -or `
            $TeamsStatus -like "*StatusIndicatorStateService: Added Offline*") {
        $Status = $lgOffline
        Write-Host $Status
    }

    If ($TeamsActivity -like "*Resuming daemon App updates*" -or `
        $TeamsActivity -like "*SfB:TeamsNoCall*") {
        $Activity = $lgNotInACall
        $ActivityIcon = $iconNotInACall
        Write-Host $Activity
    }
    ElseIf ($TeamsActivity -like "*Pausing daemon App updates*" -or `
        $TeamsActivity -like "*SfB:TeamsActiveCall*") {
        $Activity = $lgInACall
        $ActivityIcon = $iconInACall
        Write-Host $Activity
    }
}
# Set status to Offline when the Teams application is not running
Else {
        $Status = $lgOffline
        $Activity = $lgNotInACall
        $ActivityIcon = $iconNotInACall
        Write-Host $Status
        Write-Host $Activity
}

# Call Home Assistant API to set the status and activity sensors
If ($CurrentStatus -ne $Status) {
    $CurrentStatus = $Status

    $params = @{
     "state"="$CurrentStatus";
     "attributes"= @{
        "friendly_name"="Microsoft Teams status";
        "icon"="mdi:microsoft-teams";
        }
     }
	 
    $params = $params | ConvertTo-Json
    Invoke-RestMethod -Uri "$HAUrl/api/states/sensor.teams_status" -Method POST -Headers $headers -Body ([System.Text.Encoding]::UTF8.GetBytes($params)) -ContentType "application/json" 
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
    $params = $params | ConvertTo-Json
    Invoke-RestMethod -Uri "$HAUrl/api/states/sensor.teams_activity" -Method POST -Headers $headers -Body ([System.Text.Encoding]::UTF8.GetBytes($params)) -ContentType "application/json" 
}
    Start-Sleep 1
} Until ($Enable -eq 0)
