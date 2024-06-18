# Import Settings PowerShell script
. ($PSScriptRoot + "\Settings.ps1")

# Initialize global variables
$currentAvailability = $null
$currentActivity = $null
$previousAvailability = $null
$previousActivity = $null
$ActivityIcon = $iconNotInACall

function Get-LatestLogFile {
    try {
        $logFiles = Get-ChildItem -Path $logDirPath -Filter "MSTeams_*.log" | Sort-Object LastWriteTime -Descending
        if ($logFiles.Count -gt 0) {
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

Function Find-LatestAvailability {
    Param ([string]$logFilePath)

    $logFileContent = Get-Content -Path $logFilePath -ReadCount 0 | Select-Object -Last 1000
    [array]::Reverse($logFileContent)
    foreach ($line in $logFileContent) {
        $timestampMatches = [regex]::Matches($line, $timestampPattern)
        $availabilityMatches = [regex]::Matches($line, $availabilityPattern)

        if ($timestampMatches.Count -gt 0 -and $availabilityMatches.Count -gt 0) {
            $availabilitytimestamp = $timestampMatches[0].Value
            $script:currentAvailability = $availabilityMatches[0].Groups[1].Value
            Write-Host "Found availability: $script:currentAvailability at $availabilitytimestamp"
            return "Timestamp: $availabilitytimestamp, Availability: $script:currentAvailability"
        }
    }
    Write-Host "No availability found in the recent log entries."
    return $null
}

Function Find-LatestActivity {
    Param ([string]$logFilePath)

    $logFileContent = Get-Content -Path $logFilePath -ReadCount 0 | Select-Object -Last 1000
    [array]::Reverse($logFileContent)
    foreach ($line in $logFileContent) {
        $timestampMatches = [regex]::Matches($line, $timestampPattern)
        $activityMatches = [regex]::Matches($line, $activityPattern)

        if ($timestampMatches.Count -gt 0 -and $activityMatches.Count -gt 0) {
            $activitytimestamp = $timestampMatches[0].Value
            $script:currentActivity = $activityMatches[0].Groups[1].Value
            Write-Host "Found activity: $script:currentActivity at $activitytimestamp"
            return "Timestamp: $activitytimestamp, Activity: $script:currentActivity"
        }
    }
    Write-Host "No activity found in the recent log entries."
    return $null
}

Function Set-ActivityIcon {
    if ($script:currentActivity -eq "VeryActive") {
        $script:ActivityIcon = $iconInACall
    } else {
        $script:ActivityIcon = $iconNotInACall
    }
}

Function Send-LatestActivity {
    if ($script:currentAvailability -ne $script:previousAvailability) {
        $params = @{
            "state" = "$script:currentAvailability";
            "attributes" = @{
                "friendly_name" = "$entityStatusName";
                "icon" = "mdi:microsoft-teams";
            }
        }
        $params = $params | ConvertTo-Json
        try {
            Invoke-RestMethod -Uri "$HAUrl/api/states/$entityStatus" -Method POST -Headers $headers -Body ([System.Text.Encoding]::UTF8.GetBytes($params)) -ContentType "application/json"
            Write-Host "Successfully updated availability to Home Assistant."
        } catch {
            Write-Error "Failed to update availability to Home Assistant: $_"
        }
    } else {
        Write-Host "No change in availability, skipping update."
    }

    if ($script:currentActivity -ne $script:previousActivity) {
        $params = @{
            "state" = "$script:currentActivity";
            "attributes" = @{
                "friendly_name" = "$entityActivityName";
                "icon" = "$script:ActivityIcon";
            }
        }
        $params = $params | ConvertTo-Json
        try {
            Invoke-RestMethod -Uri "$HAUrl/api/states/$entityActivity" -Method POST -Headers $headers -Body ([System.Text.Encoding]::UTF8.GetBytes($params)) -ContentType "application/json"
            Write-Host "Successfully updated activity to Home Assistant."
        } catch {
            Write-Error "Failed to update activity to Home Assistant: $_"
        }
    } else {
        Write-Host "No change in activity, skipping update."
    }

    $script:previousAvailability = $script:currentAvailability
    $script:previousActivity = $script:currentActivity
}

while ($true) {
    $TeamsProcess = Get-Process -Name ms-teams -ErrorAction SilentlyContinue

    if ($null -ne $TeamsProcess) {
        Write-Host "Microsoft Teams process is running."
        
        $latestLogFile = Get-LatestLogFile
        if ($latestLogFile -ne $null) {
            $latestAvailabilityEntry = Find-LatestAvailability -logFilePath $latestLogFile
            if ($latestAvailabilityEntry -ne $null) {
                Write-Host $latestAvailabilityEntry
            } else {
                Write-Host "No availability entry found in the log file."
            }

            $latestActivityEntry = Find-LatestActivity -logFilePath $latestLogFile
            if ($latestActivityEntry -ne $null) {
                Write-Host $latestActivityEntry
            } else {
                Write-Host "No activity entry found in the log file."
            }

            Set-ActivityIcon
        } else {
            Write-Host "No log file found."
        }
    } else {
        Write-Host "Microsoft Teams process is not running."
        $script:currentAvailability = "Offline"
        $script:currentActivity = "Inactive"
        $script:ActivityIcon = $iconNotInACall
        Write-Host "Availability set to: $script:currentAvailability"
        Write-Host "Activity set to: $script:currentActivity"
    }

    Send-LatestActivity
    Start-Sleep -Seconds $refreshDelay
}
