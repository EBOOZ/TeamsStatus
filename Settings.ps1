# Configure the variables below that will be used in the script
$HAToken = "eyJ0eXAiOiJKV1QiLCJhbGciOiJIUzI1NiJ9.eyJpc3MiOiJhMTA0NjgwNjRlZTM0YmQ0OGY5NTE3M2I2NmM1ZWY3MyIsImlhdCI6MTY0MTIyNjQ1NSwiZXhwIjoxOTU2NTg2NDU1fQ.GkLeqakMmJRCaejhbt4sQKc85caXeJOUTi5B3mUx2y8" # Example: eyJ0eXAiOiJKV1...
$UserName = "pwatkin1" # When not sure, open a command prompt and type: echo %USERNAME%
$HAUrl = "https://homeassistant.narcolepsy.ninja" # Example: https://yourha.duckdns.org

# Set language variables below
$lgAvailable = "Available"
$lgBusy = "Busy"
$lgOnThePhone = "On the phone"
$lgAway = "Away"
$lgBeRightBack = "Be right back"
$lgDoNotDisturb = "Do not disturb"
$lgPresenting = "Presenting"
$lgFocusing = "Focusing"
$lgInAMeeting = "In a meeting"
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
