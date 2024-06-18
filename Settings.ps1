# Configure the variables below that will be used in the script
$HAToken = "eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJpc3MiOiJhYTgwNWQxYjZjNzA0NGUzYTFlNGI1OThjM2U5NmQ5OCIsImlhdCI6MTcxODM4NzkxMSwiZXhwIjoyMDMzNzQ3OTExfQ.azOQffUALX50S8Jqzm6ZTdBzYLEtOksyRADXkE_W0tw" # Example: eyJ0eXAiOiJKV1...
$UserName = "rmclella" # When not sure, open a command prompt and type: echo %USERNAME%
$HAUrl = "http://192.168.1.16:8123" # Example: https://yourha.duckdns.org
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
