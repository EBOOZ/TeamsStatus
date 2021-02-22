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
        "friendly_name"="$entityStatusName";
        "icon"="mdi:microsoft-teams";
        }
     }
	 
    $params = $params | ConvertTo-Json
    Invoke-RestMethod -Uri "$HAUrl/api/states/$entityStatus" -Method POST -Headers $headers -Body ([System.Text.Encoding]::UTF8.GetBytes($params)) -ContentType "application/json"
	
    break
}

# Create function that will be used to set the status of the Windows Service
Function publishOnlineState ()
{
    $params = @{
        "state"="updating";
        "attributes"= @{
           "friendly_name"="$entityHeartbeatName";
           "icon"="$iconMonitoring";
           "device_class"="connectivity";
           }
        }
   
    $params = $params | ConvertTo-Json
    Invoke-RestMethod -Uri "$HAUrl/api/states/$entityHeartbeat" -Method POST -Headers $headers -Body ([System.Text.Encoding]::UTF8.GetBytes($params)) -ContentType "application/json" 

    $params = @{
        "state"="on";
        "attributes"= @{
           "friendly_name"="$entityHeartbeatName";
           "icon"="$iconMonitoring";
           "device_class"="connectivity";
           }
        }
   
    $params = $params | ConvertTo-Json
    Invoke-RestMethod -Uri "$HAUrl/api/states/$entityHeartbeat" -Method POST -Headers $headers -Body ([System.Text.Encoding]::UTF8.GetBytes($params)) -ContentType "application/json" 
}

# Clear the status and activity variable at the start
$CurrentStatus = ""
$CurrentActivity = ""

# Start the stopwatch that will be monitoring the status of the Windows Service
$stopwatch =  [system.diagnostics.stopwatch]::StartNew()
# Set the Windows Service heartbeat sensor at startup
publishOnlineState

# Start monitoring the Teams logfile when no parameter is used to run the script
DO {
# Set the Windows Service heartbeat sensor every 5 minutes and restart the stopwatch
If ([int]$stopwatch.Elapsed.Minutes -ge 4){
    $stopwatch.Restart()
    publishOnlineState
}

# Get Teams Logfile and last icon overlay status
$TeamsStatus = Get-Content -Path "C:\Users\$UserName\AppData\Roaming\Microsoft\Teams\logs.txt" -Tail 200 | Select-String -Pattern `
  'Setting the taskbar overlay icon -',`
  'StatusIndicatorStateService: Added' | Select-Object -Last 1
Write-Host $TeamsStatus
# Get Teams Logfile and last app update deamon status
$TeamsActivity = Get-Content -Path "C:\Users\$UserName\AppData\Roaming\Microsoft\Teams\logs.txt" -Tail 200 | Select-String -Pattern `
  'Resuming daemon App updates',`
  'Pausing daemon App updates',`
  'SfB:TeamsNoCall',`
  'SfB:TeamsPendingCall',`
  'SfB:TeamsActiveCall',`
  'StatusIndicatorStateService: Added' | Select-Object -Last 1
Write-Host $TeamsActivity
# Get Teams application process
$TeamsProcess = Get-Process -Name Teams -ErrorAction SilentlyContinue

# Check if Teams is running and start monitoring the log if it is
If ($null -ne $TeamsProcess) {
    If ($TeamsStatus -like "*Setting the taskbar overlay icon - $lgAvailable*" -or `
        $TeamsStatus -like "*StatusIndicatorStateService: Added Available*") {
        $Status = $lgAvailable
        Write-Host $Status
    }
    ElseIf ($TeamsStatus -like "*Setting the taskbar overlay icon - $lgBusy*" -or `
            $TeamsStatus -like "*StatusIndicatorStateService: Added Busy*" -or `
            $TeamsStatus -like "*Setting the taskbar overlay icon - $lgOnThePhone*" -or `
            $TeamsStatus -like "*StatusIndicatorStateService: Added OnThePhone*") {
        $Status = $lgBusy
        Write-Host $Status
    }
    ElseIf ($TeamsStatus -like "*Setting the taskbar overlay icon - $lgAway*" -or `
            $TeamsStatus -like "*StatusIndicatorStateService: Added Away*") {
        $Status = $lgAway
        Write-Host $Status
    }
    ElseIf ($TeamsStatus -like "*Setting the taskbar overlay icon - $lgBeRightBack*" -or `
            $TeamsStatus -like "*StatusIndicatorStateService: Added BeRightBack*") {
        $Status = $lgBeRightBack
        Write-Host $Status
    }
    ElseIf ($TeamsStatus -like "*Setting the taskbar overlay icon - $lgDoNotDisturb *" -or `
            $TeamsStatus -like "*StatusIndicatorStateService: Added DoNotDisturb*") {
        $Status = $lgDoNotDisturb
        Write-Host $Status
    }
    ElseIf ($TeamsStatus -like "*Setting the taskbar overlay icon - $lgFocusing*" -or `
            $TeamsStatus -like "*StatusIndicatorStateService: Added Focusing*") {
        $Status = $lgFocusing
        Write-Host $Status
    }
    ElseIf ($TeamsStatus -like "*Setting the taskbar overlay icon - $lgPresenting*" -or `
            $TeamsStatus -like "*StatusIndicatorStateService: Added Presenting*") {
        $Status = $lgPresenting
        Write-Host $Status
    }
    ElseIf ($TeamsStatus -like "*Setting the taskbar overlay icon - $lgInAMeeting*" -or `
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
        $TeamsActivity -like "*SfB:TeamsNoCall*" -or `
        $TeamsActivity -like "*OnThePhone -> $lgAvailable*"-or `
        $TeamsActivity -like "*OnThePhone -> $lgBusy*"-or `
        $TeamsActivity -like "*OnThePhone -> $lgInAMeeting*"-or `
        $TeamsActivity -like "*OnThePhone -> $lgFocusing*"-or `
        $TeamsActivity -like "*OnThePhone -> $lgDoNotDisturb*" -or `
        $TeamsActivity -like "*OnThePhone -> Available*"-or `
        $TeamsActivity -like "*OnThePhone -> Busy*"-or `
        $TeamsActivity -like "*OnThePhone -> InAMeeting*"-or `
        $TeamsActivity -like "*OnThePhone -> Focusing*"-or `
        $TeamsActivity -like "*OnThePhone -> DoNotDisturb*") {
        $Activity = $lgNotInACall
        $ActivityIcon = $iconNotInACall
        Write-Host $Activity
    }
    ElseIf ($TeamsActivity -like "*Pausing daemon App updates*" -or `
        $TeamsActivity -like "*SfB:TeamsActiveCall*" -or `
        $TeamsActivity -like "*Setting the taskbar overlay icon - $lgOnThePhone*" -or `
        $TeamsActivity -like "*StatusIndicatorStateService: Added OnThePhone*" -or `
        $TeamsActivity -like "*-> OnThePhone*") {
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