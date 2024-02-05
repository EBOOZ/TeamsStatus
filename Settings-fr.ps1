# Configure the variables below that will be used in the script
$UserName = "<UserName>" # Need when running as a service (When not sure, open a command prompt and type: echo %USERNAME%)

$HAToken = "<Insert token>" # Example: eyJ0eXAiOiJKV1...
$HAUrl = "<HAUrl>" # Example: https://yourha.duckdns.org

$JeedomToken = "<Insert token>" # Example: eyJ0eXAiOiJKV1...
$JeedomUrl = "<JeedomUrl>" # Example: https://yourjeedom.duckdns.org
$JeedomStatusId = "<ID>" # Example : 745
$JeedomActivityId = "<ID>" # Example : 746

# Set language variables below
$lgAvailable = "Disponible"
$lgBusy = "Occupé"
$lgOnThePhone = "Au téléphone"
$lgAway = "Absent(e)"
$lgBeRightBack = "De retour bientôt"
$lgDoNotDisturb = "Ne pas déranger"
$lgPresenting = "En présentation"
$lgFocusing = "Focusing"
$lgInAMeeting = "En réunion"
$lgOffline = "Hors connexion"
$lgNotInACall = "Plus en appel"
$lgInACall = "En appel"

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