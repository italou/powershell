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
   Version: 1.1
   Updated: 2021/03/12

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
    Write-Host Checking module $module.name... " " -NoNewline

    $moduleVersions = Get-InstalledModule $module.name -allversions

    Write-Host Latest installed version is $moduleVersions[-1].version .

    if ($moduleVersions.count -gt 1 )
    {
        Write-Host Found $moduleVersions.count versions. -ForegroundColor Yellow

        # Iterate through versions and unsinstall previous ones.
        foreach ($moduleVersion in $moduleVersions)
        {
            if ($moduleVersion.version -ne $moduleVersions[-1].version)
            {
                Write-Host "   "Uninstalling version $moduleVersion.version ... -ForegroundColor Yellow
                $moduleVersion | Uninstall-Module -force
            }
        }
    }

    Write-Host "------------------------"
}

Write-Host "All modules checked!" -ForegroundColor Green
