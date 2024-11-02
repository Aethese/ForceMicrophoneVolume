# import ADC
$scriptDir = Split-Path -parent $MyInvocation.MyCommand.Path
try
{
	Import-Module $scriptDir\Modules\AudioDeviceCmdlets\3.1.0.2\AudioDeviceCmdlets.psd1 -ErrorAction Stop
}
catch [System.IO.FileNotFoundException]
{
	Write-Warning "AudioDeviceCmdlets not found in path. Path checked: $scriptDir\Modules\AudioDeviceCmdlets\3.1.0.2\AudioDeviceCmdlets.psd1"
	exit
}

# handle volume level
$userInput = Read-Host "Your preferred volume level (as a percentage number)"
$volumeLevel = $userInput -as [int] # convert to int

if ($volumeLevel -isnot [int])
{
	Write-Warning "Expected, but did not get, int for volume level"
	exit
}

# get mic
Write-Output "Select by Index which microphone you'd like to change:"
Start-Sleep -Seconds 1
Get-AudioDevice -List
$index = Read-Host "Index"

try
{
	$device = Get-AudioDevice -Index $index
	$deviceID = $device.ID
}
catch [System.ArgumentException]
{
	Write-Warning "No device found at index $index"
	exit
}

Set-AudioDevice -Id $deviceID # select mic
Set-AudioDevice -RecordingCommunicationVolume $volumeLevel # change volume

# run loop
$iteration = 0
$runTime = 4 # in hours
Write-Output "Setting microphone value to $volumeLevel% for $runTime hour(s). Press ctrl+c to stop the script early"

# hours -> minutes -> ms
while ($iteration -lt ($runTime * 60 / 0.5))
{
	Set-AudioDevice -RecordingCommunicationVolume $volumeLevel
	$iteration += 1
	Start-Sleep -Milliseconds 500
}

Write-Output "Stopped running after $runTime hour(s)"
