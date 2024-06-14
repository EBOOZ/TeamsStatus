<#
.NOTES
    Name: Get-TeamsStatus.ps1
    Author: Danny de Vries
    New Teams Compatibility Hack: @latetedemelon
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

# Define the function to get the latest log file
function Get-LatestLogFile {
    try {
        $logFiles = Get-ChildItem -Path $logDirPath -Filter "MSTeams_*.log" | Sort-Object LastWriteTime -Descending
        if ($logFiles.Count -gt 0) {
            Write-Host "Latest log file: $($logFiles[0].FullName)"
            return $logFiles[0].FullName
        } else {
            Write-Error "No log files found."
            return $null
        }
    } catch {
        Write-Error "Failed to fetch log files: $_"
        return $null
    }
}

# Start monitoring the Teams logfile when no parameter is used to run the script
DO {
    # Get Teams Logfile and last icon overlay status
    $latestLogFile = Get-LatestLogFile
    if ($null -eq $latestLogFile) {
        Write-Host "No log file found, skipping this iteration."
        Start-Sleep 1
        continue
    }

    try {
        $logContent = Get-Content -Path $latestLogFile -Tail 1000 | Select-Object -Last 1000
        $reversedContent = $logContent[$logContent.Count..0]
        Write-Host "Successfully read the log file content."
    } catch {
        Write-Error "Failed to read log file content: $_"
        Start-Sleep 1
        continue
    }

    $TeamsStatus = $reversedContent | Select-String -Pattern `
      'SetTaskbarIconOverlay overlay description',`
      'Received Action: UserPresenceAction:',`
      'CloudStateChanged: New Cloud State Event:',`
      'BroadcastGlobalState: New Global State Event:',`
      'Received Action: UserPresenceAction:'
    Write-Host "TeamsStatus entries found: $($TeamsStatus.Count)"

    # Get Teams Logfile and last app update daemon status
    $TeamsActivity = $reversedContent | Select-String -Pattern `
      'EmitWebClientStateChangeEvent',`
      'SaveHighestWebClientState',`
      'SaveLastChangedWebClientState'
    Write-Host "TeamsActivity entries found: $($TeamsActivity.Count)"

    # Get Teams application process
    $TeamsProcess = Get-Process -Name ms-teams -ErrorAction SilentlyContinue

    # Check if Teams is running and start monitoring the log if it is
    If ($null -ne $TeamsProcess) {
        Write-Host "Microsoft Teams process is running."
        If($TeamsStatus -eq $null){ 
            Write-Host "No TeamsStatus entries found."
        } ElseIf ($TeamsStatus -like "*availability: $lgAvailable*") {
            $Status = $lgAvailable
            Write-Host "Status set to: $Status"
        } ElseIf ($TeamsStatus -like "*availability: $lgBusy*") {
            $Status = $lgBusy
            Write-Host "Status set to: $Status"
        } ElseIf ($TeamsStatus -like "*availability: $lgAway*") {
            $Status = $lgAway
            Write-Host "Status set to: $Status"
        } ElseIf ($TeamsStatus -like "*availability: $lgBeRightBack*") {
            $Status = $lgBeRightBack
            Write-Host "Status set to: $Status"
        } ElseIf ($TeamsStatus -like "*availability: $lgDoNotDisturb*") {
            $Status = $lgDoNotDisturb
            Write-Host "Status set to: $Status"
        } ElseIf ($TeamsStatus -like "*availability: $lgFocusing*") {
            $Status = $lgFocusing
            Write-Host "Status set to: $Status"
        } ElseIf ($TeamsStatus -like "*availability: $lgPresenting*") {
            $Status = $lgPresenting
            Write-Host "Status set to: $Status"
        } ElseIf ($TeamsStatus -like "*availability: $lgInAMeeting*") {
            $Status = $lgInAMeeting
            Write-Host "Status set to: $Status"
        } ElseIf ($TeamsStatus -like "*availability: $lgOffline*") {
            $Status = $lgOffline
            Write-Host "Status set to: $Status"
        }

        If($TeamsActivity -eq $null){ 
            Write-Host "No TeamsActivity entries found."
        } ElseIf ($TeamsActivity -notlike "*state=VeryActive*") {
            $Activity = $lgNotInACall
            $ActivityIcon = $iconNotInACall
            Write-Host "Activity set to: $Activity"
        } ElseIf ($TeamsActivity -like "*state=VeryActive*") {
            $Activity = $lgInACall
            $ActivityIcon = $iconInACall
            Write-Host "Activity set to: $Activity"
        }
    } Else {
        # Set status to Offline when the Teams application is not running
        Write-Host "Microsoft Teams process is not running."
        $Status = $lgOffline
        $Activity = $lgNotInACall
        $ActivityIcon = $iconNotInACall
        Write-Host "Status set to: $Status"
        Write-Host "Activity set to: $Activity"
    }

    # Call Home Assistant API to set the status and activity sensors
    If ($CurrentStatus -ne $Status -and $Status -ne $null) {
        $CurrentStatus = $Status
        $params = @{
         "state"="$CurrentStatus";
         "attributes"= @{
            "friendly_name"="$entityStatusName";
            "icon"="mdi:microsoft-teams";
            }
         }
        $params = $params | ConvertTo-Json
        try {
            Invoke-RestMethod -Uri "$HAUrl/api/states/$entityStatus" -Method POST -Headers $headers -Body ([System.Text.Encoding]::UTF8.GetBytes($params)) -ContentType "application/json"
            Write-Host "Successfully updated status to Home Assistant."
        } catch {
            Write-Error "Failed to update status to Home Assistant: $_"
        }
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
        try {
            Invoke-RestMethod -Uri "$HAUrl/api/states/$entityActivity" -Method POST -Headers $headers -Body ([System.Text.Encoding]::UTF8.GetBytes($params)) -ContentType "application/json"
            Write-Host "Successfully updated activity to Home Assistant."
        } catch {
            Write-Error "Failed to update activity to Home Assistant: $_"
        }
    }

    Start-Sleep 1
} Until ($Enable -eq 0)
