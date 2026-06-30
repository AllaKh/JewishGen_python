<#
build_locked_installer.ps1 - build a HARDWARE-LOCKED installer (alternative to a normal build).

The packaged app will run ONLY on the machine id(s) you pass here. Use this instead of
build_installer.ps1 when you want the strongest offline protection (node-lock) - no server,
no payments, at the cost of one manual exchange of the recipient's machine id.

FLOW
  1. The recipient runs  get_machine_id.bat  on THEIR computer and sends you the printed id.
  2. You build a package locked to that id:
        .\build_locked_installer.ps1 -MachineId "<their-id>" -Version 2.0.1
  3. You send them  Output\JewishGenealogySearch-Setup-2.0.1.exe  (plus the install password
     if you also used -SetupPassword). On ANY OTHER computer the app shows a
     "locked to another computer" message and quits.

  Several machines in one package:  -MachineId "id1,id2,id3"
  Combine with -SetupPassword / -SelfDestruct exactly like build_installer.ps1.
#>
param(
    [Parameter(Mandatory = $true)][string]$MachineId,
    [string]$Version       = "1.0.0",
    [string]$SetupPassword = $env:SETUP_PASSWORD,
    [switch]$SelfDestruct,
    [switch]$SkipInstaller,
    [switch]$Clean
)
$ErrorActionPreference = "Stop"
& "$PSScriptRoot\build_installer.ps1" `
    -Version $Version -MachineId $MachineId `
    -SetupPassword $SetupPassword -SelfDestruct:$SelfDestruct `
    -SkipInstaller:$SkipInstaller -Clean:$Clean
