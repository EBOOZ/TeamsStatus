# Configure the variables below that will be used in the script
$HAToken = "" # Example: eyJ0eXAiOiJKV1...
$UserName = "" # When not sure, open a command prompt and type: echo %USERNAME%
$HAUrl = "" # Example: https://yourha.duckdns.org
# Define the path to the Microsoft Teams log directory
$logDirPath = "$env:UserProfile\AppData\Local\Packages\MSTeams_8wekyb3d8bbwe\LocalCache\Microsoft\MSTeams\Logs"
$refreshDelay = 5

# Set icons to use for call activity
$iconInACall = "mdi:phone-in-talk-outline"
$iconNotInACall = "mdi:phone-off"
$iconMonitoring = "mdi:api"

# Set entities to post to
$entityStatus = "sensor.microsoft_teams_status"
$entityStatusName = "Microsoft Teams status"
$entityActivity = "sensor.microsoft_teams_activity"
$entityActivityName = "Microsoft Teams activity"
$entityHeartbeat = "binary_sensor.teams_monitoring"
$entityHeartbeatName = "Microsoft Teams monitoring"

# Define the regex patterns for timestamp and availability
$timestampPattern = "\d{4}-\d{2}-\d{2}T\d{2}:\d{2}:\d{2}\.\d{6}-\d{2}:\d{2}"
$availabilityPattern = "availability: (\w+)(,|})"
$activityPattern = "(?:state=|new_state=)(\w+)"

$headers = @{"Authorization"="Bearer $HAToken";}
$Enable = 1
