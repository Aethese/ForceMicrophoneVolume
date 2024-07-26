$scriptDir = Split-Path -parent $MyInvocation.MyCommand.Path
Import-Module $scriptDir\Modules\AudioDeviceCmdlets\3.1.0.2\AudioDeviceCmdlets.psd1

$volumeLevel = Read-Host "Your preferred volume level (as a percentage)"

Write-Output "Select by Index which microphone you'd like to change:"
Start-Sleep -Seconds 1
Get-AudioDevice -List
$index = Read-Host "Index"

try {
	$device = Get-AudioDevice -Index $index
	$deviceID = $device.ID
} catch [System.ArgumentException] {
	Write-Warning "No device found at index $index"
	exit
}

Set-AudioDevice -Id $deviceID # select mic
Set-AudioDevice -RecordingCommunicationVolume $volumeLevel # change volume

$iteration = 0
$runTime = 4 # in hours
Write-Output "Setting microphone value to $volumeLevel% for $runTime hour(s). Press ctrl+c to stop the script early"

while ($iteration -lt ($runTime * 60 / 0.5)) {
	Set-AudioDevice -RecordingCommunicationVolume $volumeLevel
	$iteration += 1
	Start-Sleep -Milliseconds 500
}

Write-Output "Stopped running after $runTime hour(s)"