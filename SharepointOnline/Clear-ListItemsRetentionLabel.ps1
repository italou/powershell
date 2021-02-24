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
    Specifies the list/library name where the items are.

.NOTES
    Script will retrieve list items with pagination, set to 2500 items per page
    in order to avoid list view threshold limit (+ 5000 items) and will execute
    the update query in batches of 150 items to avoid large message errors.
    It can be also used with ODB personal sites, site collection url will be 
    the user personal site url (ie: personal/user_contoso_com) and the library
    name will the default Documents library.

    Requires: Sharepoint Online Client Components
    Version : 1.0
    Updated : 2021-02-24


.LINK
    https://docs.microsoft.com/en-us/microsoft-365/compliance/retention
    https://www.nuget.org/packages/Microsoft.SharePointOnline.CSOM
    https://www.microsoft.com/en-us/download/details.aspx?id=42038
#>


## Common Sharepoint Online Client Components path
$commonDLLInstallPath = "C:\Program Files\Common Files\microsoft shared\Web Server Extensions\16\ISAPI"

# check if common installation folder from .msi package.
if(Test-Path $commonDLLInstallPath)
{
    $clientDLLPath = $commonDLLInstallPath
    $clientRuntimeDLLPath = $commonDLLInstallPath
}
# check if already loaded via module, ie: PnP, .net framework.
elseif(([System.Reflection.Assembly]::LoadWithPartialName("Microsoft.SharePoint.Client").location))
{
    $clientDLLPath = (([System.Reflection.Assembly]::LoadWithPartialName("Microsoft.SharePoint.Client").location) -split '(.*)\\(.*)dll')[1]
    $clientRuntimeDLLPath = (([System.Reflection.Assembly]::LoadWithPartialName("Microsoft.SharePoint.Client.Runtime").location) -split '(.*)\\(.*)dll')[1]
}
# check if installed via nuget.
elseif(Get-Module -ListAvailable PackageManagement)
{
    $SPOcsom = (Get-Package | Where-Object { $_.Name -eq "Microsoft.SharePointOnline.CSOM" })

    if($SPOcsom) # installed via nuget.
    {
        $clientDLLPath = ($SPOcsom.Source -split '(.*)\\(.*)nupkg')[1]+"\lib\net45"
        $clientRuntimeDLLPath = ($SPOcsom.Source -split '(.*)\\(.*)nupkg')[1]+"\lib\net45"
    }
    else # not installed via nuget.
    {
        Write-Host "Sharepoint Online Client Components not found!" -ForegroundColor Yellow
        $clientDLLPath = Read-Host -Prompt "Enter 'Microsoft.SharePoint.Client.dll' and 'Microsoft.SharePoint.Client.Runtime.dll' files location"
        $clientRuntimeDLLPath = $clientDLLPath
    }

}
else # none from above, ask for folder.
{
    Write-Host "Sharepoint Online Client Components not found!" -ForegroundColor Yellow
    $clientDLLPath = Read-Host -Prompt "Enter 'Microsoft.SharePoint.Client.dll' and 'Microsoft.SharePoint.Client.Runtime.dll' files location"
    $clientRuntimeDLLPath = $clientDLLPath
}

# extra validation to see if needed DLL files are both present in paths
if(!(Test-Path "$clientDLLPath\Microsoft.SharePoint.Client.dll") -and !(Test-Path "$clientRuntimeDLLPath\Microsoft.SharePoint.Client.Runtime.dll"))
{
    Write-Host "Can't find '$clientDLLPath\Microsoft.SharePoint.Client.dll' or '$clientRuntimeDLLPath\Microsoft.SharePoint.Client.Runtime.dll'." -ForegroundColor Red
    break;
}

Add-Type -Path "$clientDLLPath\Microsoft.SharePoint.Client.dll"
Add-Type -Path "$clientRuntimeDLLPath\Microsoft.SharePoint.Client.Runtime.dll"

# username.
while((-not $userName) -or ($userName -eq ""))
{
    $userName = Read-Host -Prompt 'Enter username'
}

# user password.
while((-not $password) -or ($password.Length -eq 0))
{
    $password = Read-Host 'Enter password' -AsSecureString
}

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

$credentials = New-Object Microsoft.SharePoint.Client.SharePointOnlineCredentials($userName,$password)

$ctx = New-Object Microsoft.SharePoint.Client.ClientContext($webURL)
$ctx.Credentials = $credentials

try
{
    Write-Host
    Write-Host "Getting list... " -NoNewline
	$list = $ctx.Web.Lists.GetByTitle($listName)
	$ctx.load($list)
	$ctx.executeQuery()
    Write-Host "Done!" -ForegroundColor Green

    # number of items cleared.
    $numItemsCleared = 0

    ## view XML for camlQuery.
    $qCommand = @"
<View Scope="RecursiveAll">
    <Query>
        <OrderBy><FieldRef Name='ID' Ascending='TRUE'/></OrderBy>
    </Query>
    <RowLimit Paged="TRUE">2500</RowLimit>
</View>
"@

    # page position.
    $position = $null

    # get the all items by page.
    Do
    {
        Write-Host
        Write-Host "Getting list items... " -NoNewline
        $camlQuery = New-Object Microsoft.SharePoint.Client.CamlQuery
        $camlQuery.ListItemCollectionPosition = $position
        $camlQuery.ViewXml = $qCommand
        
        $currentList = $list.GetItems($camlQuery)
        $ctx.Load($currentList)
        $ctx.ExecuteQuery()
        Write-Host "Done!" -ForegroundColor Green

        # iterate through current list items
        foreach($item in $currentList)
	    {

            if($item.FieldValues._ComplianceTag -ne "")
            {
                # clear item retention label
                Write-Host "'$($item.FieldValues.FileLeafRef)' queued to clear label '$($item.FieldValues._ComplianceTag)'." -ForegroundColor Yellow
                $item.SetComplianceTag("", $false, $false, $false, $false);
                $item.Update();
                $numItemsCleared++
            }

            # updating by batches of items to avoid large message error.
            if(($numItemsCleared -gt 0) -and (($numItemsCleared % 150) -eq 0))
            {
                Write-Host "Updating queued items... " -NoNewline
                $ctx.ExecuteQuery()
                Write-Host "Done" -ForegroundColor Green
            }

	    }
 
        # last page position.
        $position = $currentList.ListItemCollectionPosition

    } while ($null -ne $position) 


    Write-Host
    Write-Host "A total of $numItemsCleared items were cleared in '$listName'." -ForegroundColor Green
}
catch
{
	Write-Host "Something went wrong: $($_.Exception.Message)" -Foregroundcolor Red
}