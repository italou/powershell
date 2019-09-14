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