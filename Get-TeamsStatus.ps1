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
    The status entity (sensor.teams_status by default) displays that availability 
    status of your Teams client based on the icon overlay in the taskbar on Windows. 
    The activity entity (sensor.teams_activity by default) shows if you
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

# Import Settings PowerShell script
. ($PSScriptRoot + "\Settings.ps1")

$headers = @{"Authorization"="Bearer $HAToken";}
$Enable = 1

# Run the script when a parameter is used and stop when done
If($null -ne $SetStatus){
    Write-Host ("Setting Microsoft Teams status to "+$SetStatus+":")
    $params = @{
     "state"="$SetStatus";
     "attributes"= @{
        "friendly_name"="$entityStatusName";
        "icon"="mdi:microsoft-teams";
        }
     }
	 
    $params = $params | ConvertTo-Json
    try {
        Invoke-RestMethod -Uri "$HAUrl/api/states/$entityStatus" -Method POST -Headers $headers -Body ([System.Text.Encoding]::UTF8.GetBytes($params)) -ContentType "application/json"
        Write-Host "Status set successfully."
    } catch {
        Write-Error "Failed to set status: $_"
    }
    break
}

# Start monitoring the Teams logfile when no parameter is used to run the script
DO {
# Get latest MSTeams_ logfile
$latestLogfile = Get-ChildItem -Path "C:\Users\$LocalUsername\AppData\Local\Packages\MSTeams_8wekyb3d8bbwe\LocalCache\Microsoft\MSTeams\Logs" -Name "MSTeams_*" | Select-Object -Last 1
$MSTeamsLog = Get-Content -Path "C:\Users\$LocalUsername\AppData\Local\Packages\MSTeams_8wekyb3d8bbwe\LocalCache\Microsoft\MSTeams\Logs\$latestLogfile" -Tail 100

# Get Teams Logfile and last icon overlay status
$TeamsStatus = $MSTeamsLog | Select-String -Pattern `
  'SetBadge Setting badge:',`
  'SetTaskbarIconOverlay' | Select-Object -Last 1

# Get Teams Logfile and last app update deamon status
$TeamsActivity = $MSTeamsLog | Select-String -Pattern `
  'NotifyCallActive',`
  'NotifyCallAccepted',`
  'NotifyCallEnded',`
  'reportIncomingCall' | Select-Object -Last 1

# Get Teams application process
$TeamsProcess = Get-Process -Name ms-teams -ErrorAction SilentlyContinue

# Check if Teams is running and start monitoring the log if it is
If ($null -ne $TeamsProcess) {
    If($null -eq $TeamsStatus){ }
    ElseIf ($TeamsStatus -like "*available*") {
        $Status = $lgAvailable
        Write-Host $Status
    }
    ElseIf ($TeamsStatus -like "*busy*") {
        $Status = $lgBusy
        Write-Host $Status
    }
    ElseIf ($TeamsStatus -like "*away*") {
        $Status = $lgAway
        Write-Host $Status
    }
    ElseIf ($TeamsStatus -like "*doNotDistrb*" -or `
        $TeamsStatus -like "*Do not disturb*") {
        $Status = $lgDoNotDisturb
        Write-Host $Status
    }
    # Dummy - Not tested yet
    ElseIf ($TeamsStatus -like "*focusing*") {
        $Status = $lgFocusing
        Write-Host $Status
    }
    # Dummy - Not tested yet
    ElseIf ($TeamsStatus -like "*presenting*") {
        $Status = $lgPresenting
        Write-Host $Status
    }
    # Dummy - Not tested yet
    ElseIf ($TeamsStatus -like "*inameeting*") {
        $Status = $lgInAMeeting
        Write-Host $Status
    }
    ElseIf ($TeamsStatus -like "*offline*") {
        $Status = $lgOffline
        Write-Host $Status
    }

    If($null -eq $TeamsActivity){ }
    ElseIf ($TeamsActivity -like "*reportIncomingCall*") {
        $Activity = $lgIncomingCall
        $ActivityIcon = $iconIncomingCall
        Write-Host $Activity
        Write-Host $ActivityIcon
    }
    ElseIf ($TeamsActivity -like "*NotifyCallActive*" -or `
        $TeamsActivity -like "*NotifyCallAccepted*") {
        $Activity = $lgInACall
        $ActivityIcon = $iconInACall
        Write-Host $Activity
        Write-Host $ActivityIcon
    }
    ElseIf ($TeamsActivity -like "*NotifyCallEnded*") {
        $Activity = $lgNotInACall
        $ActivityIcon = $iconNotInACall
        Write-Host $Activity
        Write-Host $ActivityIcon
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
If ($CurrentStatus -ne $Status -and $null -ne $Status) {
    $CurrentStatus = $Status

    $params = @{
     "state"="$CurrentStatus";
     "attributes"= @{
        "friendly_name"="$entityStatusName";
        "icon"="mdi:microsoft-teams";
        }
     }
	 
    $params = $params | ConvertTo-Json
    Invoke-RestMethod -Uri "$HAUrl/api/states/$entityStatus" -Method POST -Headers $headers -Body ([System.Text.Encoding]::UTF8.GetBytes($params)) -ContentType "application/json" 
}

If ($CurrentActivity -ne $Activity) {
    $CurrentActivity = $Activity

    $params = @{
     "state"="$Activity";
     "attributes"= @{
        "friendly_name"="$entityActivityName";
        "icon"="$ActivityIcon";
        }
     }
    $params = $params | ConvertTo-Json
    Invoke-RestMethod -Uri "$HAUrl/api/states/$entityActivity" -Method POST -Headers $headers -Body ([System.Text.Encoding]::UTF8.GetBytes($params)) -ContentType "application/json" 
}
    Start-Sleep 1
} Until ($Enable -eq 0)
