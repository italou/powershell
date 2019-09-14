<#
.Synopsis
   Checks powershell installed modules and removes old versions of them.

   ## DISCLAIMER:
   ## Copyright (c) Itamar Lourenço. All rights reserved. This
   ## script is made available to you without any express, implied or
   ## statutory warranty, not even the implied warranty of
   ## merchantability or fitness for a particular purpose, or the
   ## warranty of title or non-infringement. The entire risk of the
   ## use or the results from the use of this script remains with you.

.NOTES
   Author: Itamar Lourenço
   Created: 2019/09/13
   Version: 1.0   
#>

# Elevate to Admin mone
$currentUser = New-Object Security.Principal.WindowsPrincipal $([Security.Principal.WindowsIdentity]::GetCurrent())
$testadmin = $currentUser.IsInRole([Security.Principal.WindowsBuiltinRole]::Administrator)
if ($testadmin -eq $false) {
    Start-Process powershell.exe -Verb RunAs -ArgumentList ('-noprofile -noexit -file "{0}" -elevated' -f ($myinvocation.MyCommand.Definition))
    Write-Output "Couldn't run in Admin mode! Try manually!"
    exit $LASTEXITCODE
}

# Get installed modules
$mods = Get-InstalledModule

# Iterate through installed modules
foreach ($mod in $mods)
{
  Write-Host "Checking module: $($mod.name)"
  $latest = Get-InstalledModule $mod.name
  $specificmods = Get-InstalledModule $mod.name -allversions

  if ($specificmods.count -gt 1 )
  {
    Write-Host "Found $($specificmods.count) versions, " -NoNewline
  }

  Write-Host "latest installed version is: $($latest.version)"
  
  # Iterate through versions and unsinstall previous ones.
  foreach ($sm in $specificmods)
  {
    if ($sm.version -ne $latest.version)
	{
	  Write-Host "Uninstalling version: $($sm.version)"
	  $sm | Uninstall-Module -force
	}
	
  }
  Write-Host "------------------------"
}

Write-Host "All done!"
