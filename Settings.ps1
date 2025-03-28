# Configure the variables below that will be used in the script
$HAToken = "<Insert token>" # Example: eyJ0eXAiOiJKV1...
$HAUrl = "<HA URL>" # Example: https://yourha.duckdns.org
$LocalUsername = "<UserName>" # Example: your.username

# Set language variables below
$lgAvailable = "Availabble"
$lgBusy = "Busy"
$lgOnThePhone = "In a call"
$lgAway = "Away"
$lgBeRightBack = "Be right back"
$lgDoNotDisturb = "Do not disturb"
$lgPresenting = "Presenting"
$lgFocusing = "Focusing"
$lgInAMeeting = "In a meeting"
$lgOffline = "Offline"
$lgNotInACall = "Not in a call"
$lgInACall = "In a call"
$lgIncomingCall = "Incoming call"

# Set icons to use for call activity
$iconInACall = "mdi:phone-in-talk-outline"
$iconNotInACall = "mdi:phone-off"
$iconIncomingCall = "mdi:phone-ring"

# Set entities to post to
$entityStatus = "sensor.teams_status"
$entityStatusName = "Microsoft Teams status"
$entityActivity = "sensor.teams_activity"
$entityActivityName = "Microsoft Teams activity"
