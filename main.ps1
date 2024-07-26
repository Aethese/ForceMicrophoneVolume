$scriptDir = Split-Path -parent $MyInvocation.MyCommand.Path
Import-Module $scriptDir\Modules\AudioDeviceCmdlets\3.1.0.2\AudioDeviceCmdlets.psd1

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
Set-AudioDevice -RecordingCommunicationVolume 100 # change volume

Write-Output "Setting microphone value to 100%. Press ctrl+c to stop the script"

$iteration = 0
$runTime = 4 # in hours
while ($iteration -lt ($runTime * 60 / 0.5)) {
	Set-AudioDevice -RecordingCommunicationVolume 100 # change volume
	$iteration += 1
	Start-Sleep -Milliseconds 500
}

Write-Output "Stopped running after $runTime hour(s)"