# Configure the variables below that will be used in the script
$HAToken = "" # Example: eyJ0eXAiOiJKV1...
$UserName = "" # When not sure, open a command prompt and type: echo %USERNAME%
$HAUrl = "" # Example: https://yourha.duckdns.org
# Define the path to the Microsoft Teams log directory
$logDirPath = "$env:UserProfile\AppData\Local\Packages\MSTeams_8wekyb3d8bbwe\LocalCache\Microsoft\MSTeams\Logs"
$refreshDelay = 5

# Set language variables below
$lgAvailable = "Available"
$lgBusy = "Busy"
$lgOnThePhone = "Onthephone"
$lgAway = "Away"
$lgBeRightBack = "BeRightBack"
$lgDoNotDisturb = "DoNotDisturb"
$lgPresenting = "Presenting"
$lgFocusing = "Focusing"
$lgInAMeeting = "InAMeeting"
$lgOffline = "Offline"
$lgNotInACall = "Not in a call"
$lgInACall = "In a call"

# Set icons to use for call activity
$iconInACall = "mdi:phone-in-talk-outline"
$iconNotInACall = "mdi:phone-off"
$iconMonitoring = "mdi:api"

# Set entities to post to
$entityStatus = "sensor.teams_status"
$entityStatusName = "Microsoft Teams status"
$entityActivity = "sensor.teams_activity"
$entityActivityName = "Microsoft Teams activity"
$entityHeartbeat = "binary_sensor.teams_monitoring"
$entityHeartbeatName = "Microsoft Teams monitoring"
