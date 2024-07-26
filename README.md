# Force Microphone Volume
Forces your global microphone volume on Windows to stay at your desired volume level. Helpful when apps force automatic voice control with no way of turning it off.

## Installation
You will need to install the AudioDeviceCmdlets module by running `Install-Module -Name AudioDeviceCmdlets` as administrator.

I personally placed the module inside of this directory to fix some issues I was having. If you don't do this approach, you will need to change the import-module path, possible to just `Import-Module AudioDeviceCmdlets`

## Running Script
**Security Disclaimer:** Running unsigned scripts poses a security risk! Please make sure any files you intent to give elevated permissions have been verified by you or someone you trust to be safe!

By default, you most likely won't be able to run the script directly. You may need to run a command that gives it temporary elevated permissions.

#### Run with elevated permissions
1. Open Powershell as admin and open the path that the project is held in
2. Run the command below to give the script temporary eleveted permissions (read *Security Disclaimer* above first)
```
powershell -ExecutionPolicy Bypass -File main.ps1
```

#### Run normally from Powershell
1. Open Powershell as admin and open the path that you the project is held in
2. Run the Powershell script by running `./main.ps1`