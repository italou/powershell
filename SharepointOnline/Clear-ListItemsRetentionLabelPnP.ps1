<#
    ## DISCLAIMER:
    ## The sample scripts provided here are not supported under any Microsoft
    ## standard support program or service, neither by the Author(s) and
    ## Contributor(s). All scripts are provided AS IS without warranty of any
    ## kind. Microsoft, the Author(s) and the Contributor(s) further disclaims
    ## all implied warranties including, without limitation, any implied
    ## warranties of merchantability or of fitness for a particular purpose.
    ## The entire risk arising out of the use or performance of the sample
    ## scripts and documentation remains with you. In no event shall Microsoft,
    ## its Author(s), or anyone else involved in the creation, production, or
    ## delivery of the scripts be liable for any damages whatsoever (including,
    ## without limitation, damages for loss of business profits, business
    ## interruption, loss of business information, or other pecuniary loss)
    ## arising out of the use of or inability to use the sample scripts or 
    ## documentation, even if Microsoft has been advised of the possibility of
    ## such damages.


.SYNOPSIS  
    Clears the retention label from all library items.

.DESCRIPTION
    This script will iterate through all of the items in the specified library
    and will clear the retention label applied to the items (compliance tag).

.PARAMETER webURL
    Specifies the site collection url where the items are.

.PARAMETER listName
    Specifies the list (library) name where the items are.


.NOTES
    Requires: PnP PowerShell
    Version : 1.01
    Updated : 2023-08-22


.LINK
    https://docs.microsoft.com/en-us/microsoft-365/compliance/retention
    https://docs.microsoft.com/en-us/powershell/sharepoint/sharepoint-pnp/sharepoint-pnp-cmdlets

#>


# Check if (non legacy module) PnP module is installed, if not, warn and break
if(!(Get-Module -ListAvailable -Name PnP.PowerShell)) { Write-Host "PnP.PowerShell module not found!" -ForegroundColor Red; break;}

# site collection url.
while((-not $webURL) -or ($webURL -eq ""))
{
    $webURL = Read-Host 'Enter site collection URL'
}

# list name.
while((-not $listName) -or ($listName -eq ""))
{
    $listName = Read-Host 'Enter list/library name'
}

# number of items cleared.
$numItemsCleared = 0

try
{
    Connect-PnPOnline -Url $webURL -Interactive

    Write-Host
    Write-Host "Getting list items... " -NoNewline
    $itemsList = Get-PnPListItem -List $listName -PageSize 2500
    Write-Host "Done!" -ForegroundColor Green

    # iterate through list items
    foreach($item in $itemsList)
    {
        if($item.FieldValues._ComplianceTag -ne "")
        {

            Write-Host "Clearing label '$($item.FieldValues._ComplianceTag)' on item '$($item.FieldValues.FileLeafRef)'... " -ForegroundColor Yellow -NoNewline

            # clear item retention label
            $result = Set-PnPListItem -List $listName -Identity $item.FieldValues.ID -ClearLabel
            
            Write-Host "Cleared!" -ForegroundColor Green
            $numItemsCleared++
        }
    }

    Write-Host
    Write-Host "A total of $numItemsCleared items were cleared in $listName." -ForegroundColor Green

}
catch
{
    Write-Host "Something went wrong: $($_.Exception.Message)" -Foregroundcolor Red
}
