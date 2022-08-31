<#
.NOTES
    Name: Get-TeamsStatus.ps1
    Author: Danny de Vries
    Requires: PowerShell v2 or higher
    Version History: https://github.com/EBOOZ/TeamsStatus/commits/main
.SYNOPSIS
    Sets the status of the Microsoft Teams client to Home Assistant/Jeedom.
.DESCRIPTION
    This script is monitoring the Teams client logfile for certain changes.
    For Home Assistant :
    It makes use of two sensors that are created in Home Assistant up front.
    The status entity (sensor.teams_status by default) displays that availability 
    status of your Teams client based on the icon overlay in the taskbar on Windows. 
    The activity entity (sensor.teams_activity by default) shows if you
    are in a call or not based on the App updates deamon, which is paused as soon as 
    you join a call.
    For Jeedom :
    You need to create a Virtual with 2 sensors and type "other" and get their ID.
    One for the status and one for the activity.
.PARAMETER SetStatus
    Run the script with the SetStatus-parameter to set the status of Microsoft Teams
    directly from the commandline.
.EXAMPLE
    .\Get-TeamsStatus.ps1 -SetStatus "Offline"
#>
# Configuring parameter for interactive run
Param($SetStatus)

function Convert-Umlaut
{
  param
  (
    [Parameter(Mandatory)]
    $Text
  )
       
  $output = $Text.Replace('Ã©','é').Replace('Ã´','ô')
  $isCapitalLetter = $Text -ceq $Text.toUpper()
  if ($isCapitalLetter) 
  { 
    $output = $output.toUpper() 
  }
  $output
}

# Import Settings PowerShell script
. ($PSScriptRoot + "\Settings.local.ps1")

if (($null -ne $HAToken) -and ($null -ne $HAUrl)) {
    $headers = @{"Authorization"="Bearer $HAToken";}
}

$Enable = 1

# Run the script when a parameter is used and stop when done
If($null -ne $SetStatus){
    Write-Host ("Setting Microsoft Teams status to "+$SetStatus+":")
    
    if (($null -ne $HAToken) -and ($null -ne $HAUrl)) {
        $params = @{
        "state"="$SetStatus";
        "attributes"= @{
            "friendly_name"="$entityStatusName";
            "icon"="mdi:microsoft-teams";
            }
        }    
        $params = $params | ConvertTo-Json
        Invoke-RestMethod -Uri "$HAUrl/api/states/$entityStatus" -Method POST -Headers $headers -Body ([System.Text.Encoding]::UTF8.GetBytes($params)) -ContentType "application/json"
    }
    if (($null -ne $JeedomToken) -and ($null -ne $JeedomUrl)) {
        Invoke-WebRequest -UseBasicParsing -Uri "https://$JeedomUrl/core/api/jeeApi.php?plugin=virtual&type=event&apikey=$JeedomToken&id=$JeedomStatusId&value=$SetStatus"
    }
    break
}

# Start monitoring the Teams logfile when no parameter is used to run the script
DO {
# Get Teams Logfile and last icon overlay status
$TeamsStatus = Get-Content -Path $env:APPDATA"\Microsoft\Teams\logs.txt" -encoding utf8 -Tail 1000 | Select-String -Pattern `
  'Setting the taskbar overlay icon -',`
  'StatusIndicatorStateService: Added' | Select-Object -Last 1 

# To debug another language
# Write-Host "test variable UTF8 $(Convert-Umlaut($lgAvailable)) $(Convert-Umlaut($lgBusy)) $(Convert-Umlaut($lgOnThePhone)) $(Convert-Umlaut($lgAway)) $(Convert-Umlaut($lgBeRightBack)) $(Convert-Umlaut($lgDoNotDisturb)) $(Convert-Umlaut($lgPresenting)) $(Convert-Umlaut($lgFocusing)) $(Convert-Umlaut($lgInAMeeting)) $(Convert-Umlaut($lgOffline)) $(Convert-Umlaut($lgNotInACall)) $(Convert-Umlaut($lgInACall))"

# Get Teams Logfile and last app update deamon status
$TeamsActivity = Get-Content -Path $env:APPDATA"\Microsoft\Teams\logs.txt" -encoding utf8 -Tail 1000 | Select-String -Pattern `
  'Resuming daemon App updates',`
  'Pausing daemon App updates',`
  'SfB:TeamsNoCall',`
  'SfB:TeamsPendingCall',`
  'SfB:TeamsActiveCall',`
  'name: desktop_call_state_change_send, isOngoing' | Select-Object -Last 1 

# Get Teams application process
$TeamsProcess = Get-Process -Name Teams -ErrorAction SilentlyContinue

# Check if Teams is running and start monitoring the log if it is
If ($null -ne $TeamsProcess) {
    If($null -eq $TeamsStatus){ }
    ElseIf ($TeamsStatus -like "*Setting the taskbar overlay icon - $(Convert-Umlaut($lgAvailable))*" -or `
        $TeamsStatus -like "*StatusIndicatorStateService: Added Available*" -or `
        $TeamsStatus -like "*StatusIndicatorStateService: Added NewActivity (current state: Available -> NewActivity*") {
        $Status = $(Convert-Umlaut($lgAvailable))
        Write-Host $Status
    }
    ElseIf ($TeamsStatus -like "*Setting the taskbar overlay icon - $(Convert-Umlaut($lgBusy))*" -or `
            $TeamsStatus -like "*StatusIndicatorStateService: Added Busy*" -or `
            $TeamsStatus -like "*Setting the taskbar overlay icon - $lgOnThePhone*" -or `
            $TeamsStatus -like "*StatusIndicatorStateService: Added OnThePhone*" -or `
            $TeamsStatus -like "*StatusIndicatorStateService: Added NewActivity (current state: Busy -> NewActivity*") {
        $Status = $(Convert-Umlaut($lgBusy))
        Write-Host $Status
    }
    ElseIf ($TeamsStatus -like "*Setting the taskbar overlay icon - $(Convert-Umlaut($lgAway))*" -or `
            $TeamsStatus -like "*StatusIndicatorStateService: Added Away*" -or `
            $TeamsStatus -like "*StatusIndicatorStateService: Added NewActivity (current state: Away -> NewActivity*") {
        $Status = $(Convert-Umlaut($lgAway))
        Write-Host $Status
    }
    ElseIf ($TeamsStatus -like "*Setting the taskbar overlay icon - $(Convert-Umlaut($lgBeRightBack))*" -or `
            $TeamsStatus -like "*StatusIndicatorStateService: Added BeRightBack*" -or `
            $TeamsStatus -like "*StatusIndicatorStateService: Added NewActivity (current state: BeRightBack -> NewActivity*") {
        $Status = $(Convert-Umlaut($lgBeRightBack))
        Write-Host $Status
    }
    ElseIf ($TeamsStatus -like "*Setting the taskbar overlay icon - $(Convert-Umlaut($lgDoNotDisturb)) *" -or `
            $TeamsStatus -like "*StatusIndicatorStateService: Added DoNotDisturb*" -or `
            $TeamsStatus -like "*StatusIndicatorStateService: Added NewActivity (current state: DoNotDisturb -> NewActivity*") {
        $Status = $(Convert-Umlaut($lgDoNotDisturb))
        Write-Host $Status
    }
    ElseIf ($TeamsStatus -like "*Setting the taskbar overlay icon - $(Convert-Umlaut($lgFocusing))*" -or `
            $TeamsStatus -like "*StatusIndicatorStateService: Added Focusing*" -or `
            $TeamsStatus -like "*StatusIndicatorStateService: Added NewActivity (current state: Focusing -> NewActivity*") {
        $Status = $(Convert-Umlaut($lgFocusing))
        Write-Host $Status
    }
    ElseIf ($TeamsStatus -like "*Setting the taskbar overlay icon - $(Convert-Umlaut($lgPresenting))*" -or `
            $TeamsStatus -like "*StatusIndicatorStateService: Added Presenting*" -or `
            $TeamsStatus -like "*StatusIndicatorStateService: Added NewActivity (current state: Presenting -> NewActivity*") {
        $Status = $(Convert-Umlaut($lgPresenting))
        Write-Host $Status
    }
    ElseIf ($TeamsStatus -like "*Setting the taskbar overlay icon - $(Convert-Umlaut($lgInAMeeting))*" -or `
            $TeamsStatus -like "*StatusIndicatorStateService: Added InAMeeting*" -or `
            $TeamsStatus -like "*StatusIndicatorStateService: Added NewActivity (current state: InAMeeting -> NewActivity*") {
        $Status = $(Convert-Umlaut($lgInAMeeting))
        Write-Host $Status
    }
    ElseIf ($TeamsStatus -like "*Setting the taskbar overlay icon - $(Convert-Umlaut($lgOffline))*" -or `
            $TeamsStatus -like "*StatusIndicatorStateService: Added Offline*") {
        $Status = $(Convert-Umlaut($lgOffline))
        Write-Host $Status
    }

    If($null -eq $TeamsActivity){ }
    ElseIf ($TeamsActivity -like "*Resuming daemon App updates*" -or `
        $TeamsActivity -like "*SfB:TeamsNoCall*" -or `
        $TeamsActivity -like "*name: desktop_call_state_change_send, isOngoing: false*") {
        $Activity = $(Convert-Umlaut($lgNotInACall))
        $ActivityIcon = $iconNotInACall
        Write-Host $Activity
    }
    ElseIf ($TeamsActivity -like "*Pausing daemon App updates*" -or `
        $TeamsActivity -like "*SfB:TeamsActiveCall*" -or `
        $TeamsActivity -like "*name: desktop_call_state_change_send, isOngoing: true*") {
        $Activity = $(Convert-Umlaut($lgInACall))
        $ActivityIcon = $iconInACall
        Write-Host $Activity
    }
}
# Set status to Offline when the Teams application is not running
Else {
        $Status = $(Convert-Umlaut($lgOffline))
        $Activity = $(Convert-Umlaut($lgNotInACall))
        $ActivityIcon = $iconNotInACall
        Write-Host $Status
        Write-Host $Activity
}


If ($CurrentStatus -ne $Status -and $null -ne $Status) {
    $CurrentStatus = $Status
    
    # Call Home Assistant API to set the status and activity sensors
    if (($null -ne $HAToken) -and ($null -ne $HAUrl)) {
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
    # Call Jeedom to set the status and activity sensors
    if (($null -ne $JeedomToken) -and ($null -ne $JeedomUrl)) {
        $Response = Invoke-WebRequest -Uri "https://$JeedomUrl/core/api/jeeApi.php?plugin=virtual&type=event&apikey=$JeedomToken&id=$JeedomStatusId&value=$Status"
        if ($Response|Where-Object {$_.StatusCode -ne 200}) {
            Write-Host "Error to contact Jeedom"
        }
    }
}

If ($CurrentActivity -ne $Activity) {
    $CurrentActivity = $Activity

    if (($null -ne $HAToken) -and ($null -ne $HAUrl)) {
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
    if (($null -ne $JeedomToken) -and ($null -ne $JeedomUrl)) {
        $Response = Invoke-WebRequest -Uri "https://$JeedomUrl/core/api/jeeApi.php?plugin=virtual&type=event&apikey=$JeedomToken&id=$JeedomActivityId&value=$Activity"
        if ($Response|Where-Object {$_.StatusCode -ne 200}) {
            Write-Host "Error to contact Jeedom"
        }
    }
}
# Write-Host ("status : $TeamsStatus and activity $TeamsActivity")
# Write-Host ("status $Status")
# Write-Host ("currentstatus : $CurrentStatus and currentactivity $CurrentActivity")
    Start-Sleep 1
} Until ($Enable -eq 0)
