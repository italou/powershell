<#
   ## DISCLAIMER:
   ## The script is made available to you without any express, implied or
   ## statutory warranty, not even the implied warranty of merchantability
   ## or fitness for a particular purpose, or the warranty of title or 
   ## non-infringement. The entire risk of the use or the results from the 
   ## use of this script remains with you. Please make sure to review all 
   ## script code and comments!


.SYNOPSIS  
    Script to uninstall old modules versions.
  

.DESCRIPTION
    This script will check installed modules and will uninstall all older
    versions for modules that have multiple versions installed.


.NOTES
    Author: Itamar Lourenço
   Created: 2019/09/13
   Version: 1.0

#>

# Elevate to Admin mode
$currentUser = New-Object Security.Principal.WindowsPrincipal $([Security.Principal.WindowsIdentity]::GetCurrent())
$runAsAdmin = $currentUser.IsInRole([Security.Principal.WindowsBuiltinRole]::Administrator)

if (-not $runAsAdmin)
{
    Start-Process powershell.exe -Verb RunAs -ArgumentList ('-noprofile -noexit -file "{0}" -elevated' -f ($myinvocation.MyCommand.Definition))
    Write-Output "Couldn't run in Admin mode! Try manually!"
    exit $LASTEXITCODE
}

# Get installed modules
$installedModules = Get-InstalledModule

# Iterate through installed modules
foreach ($module in $installedModules)
{
    Write-Host Checking module $module.name...

    $latestModuleVersion = Get-InstalledModule $module.name
    $allModuleVersions = Get-InstalledModule $module.name -allversions

    if ($allModuleVersions.count -gt 1 )
    {
        Write-Host Found $allModuleVersions.count versions." " -NoNewline -ForegroundColor Yellow
    }

    Write-Host Latest installed version is $latestModuleVersion.version
  
    # Iterate through versions and unsinstall previous ones.
    foreach ($moduleVersion in $allModuleVersions)
    {
        if ($moduleVersion.version -ne $latestModuleVersion.version)
        {
            Write-Host Uninstalling version $moduleVersion.version -ForegroundColor Yellow
            $moduleVersion | Uninstall-Module -force
        }
	
    }

    Write-Host "------------------------"
}

Write-Host "All modules checked!" -ForegroundColor Green
