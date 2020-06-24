#	Launcher Script for Dispatcher.
#	Version: 1.0
#	Created: 02/05/2020
#	Author: Driptaran Sen

New-Item -ItemType Directory -Force -Path "$PSScriptRoot\cache" | Out-Null
If(![System.IO.File]::Exists("$PSScriptRoot\cache\team.txt"))
{
	New-Item -path "$PSScriptRoot\cache" -name "team.txt" -type "file" | Out-Null
	$team=Read-Host "Please enter your team name(eg. Networking)"
	Set-Content -Path "$PSScriptRoot\cache\team.txt" -Value $team
	$new = 1
}
$tdata = iwr -Uri "https://address.for.json.file.with.details/Teaminfo.txt" -Method GET -UseDefaultCredentials

$data = $tdata.Content | ConvertFrom-Json

if($new -ne 1) { $team=Get-Content "$PSScriptRoot\cache\team.txt" }

if($data.$team.dispatcher -eq $null)
{
	Write-Host -ForeGroundColor Red "Sorry, team not found or supported!!!"
	rm -Force "$PSScriptRoot\cache\team.txt"
	Read-Host
	exit
}

$source = $data.$team.dispatcher
$tmp = iwr -Uri $source -Method GET -UseDefaultCredentials
$init = $tmp.Content

Invoke-Expression $init