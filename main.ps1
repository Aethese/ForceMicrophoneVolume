$scriptDir = Split-Path -parent $MyInvocation.MyCommand.Path
Import-Module $scriptDir\Modules\AudioDeviceCmdlets\3.1.0.2\AudioDeviceCmdlets.psd1

#$micID = "{0.0.1.00000000}.{5498d77b-7301-486e-8168-8c43a626ef60}"

Write-Output "Select by Index which microphone you'd like to change:"
Start-Sleep 1
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

Write-Output "Set microphone value to 100%"