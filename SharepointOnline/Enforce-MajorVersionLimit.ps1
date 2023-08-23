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
    Enforces major version limit on all library files.


.DESCRIPTION
    This script iterates through all the libray items and for files it checks 
    if the file has more versions than the set major versions limit.
    If so, it removes the oldest versions until the file only has the same 
    number of versions as the set limit.


.PARAMETER siteUrl
    Site collection url.


.PARAMETER libName
    Library name.


.NOTES
    Requires: PnP.PowerShell module
    Version : 1.01
    Updated : 2023-08-22


.LINK
    https://support.microsoft.com/en-us/office/how-versioning-works-in-lists-and-libraries-0f6cd105-974f-44a4-aadb-43ac5bdfd247

#>


# Check if (non legacy module) PnP module is installed, if not, warn and break
if(!(Get-Module -ListAvailable -Name PnP.PowerShell)) { Write-Host "PnP.PowerShell module not found!" -ForegroundColor Red; break;}

# Get site collection url
while((-not $siteUrl) -or ($siteUrl -eq ""))
{
    $siteUrl = Read-Host 'Enter site collection URL'
}

# Get library name.
while((-not $libName) -or ($libName -eq ""))
{
    $libName = Read-Host 'Enter library name'
}


try
{
    # Connect to site
    Connect-PnPOnline -Url $siteUrl -Interactive


    # Get library, for major version limit
    $lib = Get-PnPList -Identity $libName
    Write-Host "$($libName) has limit of $($lib.MajorVersionLimit) major versions." -ForegroundColor Green

    # Get list items
    Write-Host "Getting library items... "
    $items = Get-PnPListItem -List $libName -PageSize 2500

    Write-Host "Processing library items... "
    foreach($item in $items)
    {
	    if($item.FieldValues.FSObjType -eq 0) # 0 is file, 1 is folder
	    {
		    Write-Host "Checking file $($item.FieldValues.FileRef)... " -NoNewline
		
		    # Get file versions
		    $fileVersions = Get-PnPFileVersion -Url $item.FieldValues.FileRef

		    # if file has more versions than library limit
		    if($lib.MajorVersionLimit -lt $fileVersions.Count)
		    {
			    Write-Host "$($fileVersions.Count) versions found!" -ForegroundColor Red
			    $versionsToRemove = $fileVersions.Count - $lib.MajorVersionLimit

			    Write-Host "   Removing $($versionsToRemove) versions. " -NoNewline -ForegroundColor Red
	
			    for ($counter = 0; $counter -lt $versionsToRemove; $counter++)
			    {
				    Remove-PnPFileVersion -Url $item.FieldValues.FileRef -Identity $fileVersions[$counter].Id -Force
			    }
			
			    $fileVersions = Get-PnPFileVersion -Url $item.FieldValues.FileRef
			    Write-Host "File now only has $($fileVersions.Count) versions." -ForegroundColor Green
		    }
		    else
		    {
			    Write-Host "$($fileVersions.Count) versions found." -ForegroundColor Green
		    }
	    }
    }
}
catch
{
    Write-Host "Something went wrong: $($_.Exception.Message)" -Foregroundcolor Red
}
