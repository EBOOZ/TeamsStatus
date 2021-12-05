Set-ExecutionPolicy -ExecutionPolicy RemoteSigned
Unblock-File .\Settings.ps1
Unblock-File .\Lang-*.ps1
Unblock-File .\Get-TeamsStatus.ps1
Start-Process -FilePath .\nssm.exe -ArgumentList 'install "Microsoft Teams Status Monitor" "C:\Windows\System32\WindowsPowerShell\v1.0\powershell.exe" "-command "& { . C:\Scripts\Get-TeamsStatus.ps1 }"" ' -NoNewWindow -Wait
Start-Service -Name "Microsoft Teams Status Monitor"