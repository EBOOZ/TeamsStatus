> [!NOTE]
> This solution only works for the new version of Microsoft Teams. You may want to consider [Get-MonitorStatus](https://github.com/EBOOZ/Get-MonitorStatus) in combination with this script as well.

# Introduction
We're working a lot at our home office these days. Several people already found solutions to make working in the home office more comfortable. One of these ways is to automate activities in your home automatation system based on your status on Microsoft Teams.

Microsoft provides the status of your account that is used in Teams via the Graph API. To access the Graph API, your organization needs to grant consent for the organization so everybody can read their Teams status. Since my organization didn't want to grant consent, I needed to find a workaround, which I found in monitoring the Teams client logfile for certain changes.

This script makes use of two sensors that are created in Home Assistant up front:
* sensor.teams_status
* sensor.teams_activity

`sensor.teams_status` displays that availability status of your Teams client based on the icon overlay in the taskbar on Windows. `sensor.teams_activity` shows if you are in a call or not.

# Important
This solution is created to work with Home Assistant. It will work with any home automation platform that provides an API, but you probably need to change the PowerShell code.

# Requirements
* Create the two Teams sensors in the Home Assistant configuration.yaml file
```yaml
sensor:
  - platform: template
    sensors:
      teams_status: 
        friendly_name: "Microsoft Teams status"
        unique_id: sensor.teams_status
      teams_activity:
        friendly_name: "Microsoft Teams activity"
        unique_id: sensor.teams_activity
```
* Generate a Long-lived access token via `https://<HA URL>/profile/security` ([see HA documentation](https://developers.home-assistant.io/docs/auth_api/#long-lived-access-token))
* Copy and temporarily save the token somewhere you can find it later
* Restart Home Assistant to have the new sensors added
* Download the files from this repository and save them to C:\Scripts
* Edit the Settings.ps1 file and:
  * Replace `<Insert token>` with the token you generated
  * Replace `<UserName>` with the username that is logged in to Teams and you want to monitor
  * Replace `<HA URL>` with the URL to your Home Assistant server
  * Adjust the language settings to your preferences
* Start a elevated PowerShell prompt, browse to C:\Scripts and run the following command:
```powershell
Set-ExecutionPolicy -ExecutionPolicy RemoteSigned
Unblock-File .\Settings.ps1
Unblock-File .\Get-TeamsStatus.ps1
Start-Process -FilePath .\nssm.exe -ArgumentList 'install "Microsoft Teams Status Monitor" "C:\Windows\System32\WindowsPowerShell\v1.0\powershell.exe" "-command "& { . C:\Scripts\TeamsStatus\Get-TeamsStatus.ps1 }"" ' -NoNewWindow -Wait
Start-Service -Name "Microsoft Teams Status Monitor"
```

After completing the steps below, start your Teams client and verify if the status and activity is updated as expected.
