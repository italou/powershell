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
    Resolves files already existing in Sharepoint in native mode migration.
   
.DESCRIPTION
    Using the native mode migration report log file, this script will check
    for files identified as already existing in Sharepoint and append the file
    id on those so every file is migrated to Sharepoint.
    In Yammer it's possible to have different files with same name, only the
    file ids will be different, so this way we can try to migrate all files.

    To use this script it will be required to have an Authorization Token, to
    be able to perform the necessary operations in Yammer and the native mode 
    migration report log CSV file.

    The Authorization Token can be obtained by registering a Yammer App
    (see https://learn.microsoft.com/en-us/rest/api/yammer/app-registration)
    and on the App page "Generate a developer token for this application".

    The native mode migration report log file can be obtained from the native 
    mode migration page directly.

    WARNING: DO NOT EDIT AND SAVE THE FILE IN EXCEL, DOING SO CAN CHANGE THE 
             CSV ENCODING AND AFFECT SCRIPT FUNCIONALITY.


    In the end, a CSV log file will be provided with the results.
    Example: "<SCRIPTNAME>_Results_<TIMESTAMP>.csv"


.PARAMETER authToken
    Specifies the Authorization Token needed to perform the request.

.PARAMETER migrationLogCsv
    Specifies the native mode migration report log CSV filename.

.NOTES
    Version   : 1.01
    Updated   : 2023-02-07

.LINK  
    https://learn.microsoft.com/en-us/yammer/troubleshoot-problems/troubleshoot-native-mode

#>


# Warnings output
Write-Host
Write-Host "ATTENTION: You should use an Authorization Token from a GA user or Yammer App / Developer Token." -ForegroundColor Yellow
Write-Host "           Also be aware that this script was created to be used along with the native mode migration" -ForegroundColor Yellow
Write-Host "           report log CSV file provided by Yammer." -ForegroundColor Yellow
Write-Host
Write-Host "Press CTRL + C to cancel. " -NoNewline; Pause
Write-Host

# Authorization token.
[string]$authToken = ""

# Native mode migration report log file.
[string]$migrationLogCsv = ""


### Auxiliary functions ############################################

## Get file name with provided id.
function Get-YammerFilename {

    Param (
        [parameter(Mandatory=$true)][string]$FileId
    )
    
    $fileRequest = Invoke-WebRequest -Method Get `
                                     -Uri "https://www.yammer.com/api/v1/uploaded_files/$FileId.json" `
                                     -Headers @{Authorization = "Bearer $authToken"}

    $fileObject = ConvertFrom-JSON  $fileRequest.Content

    $fileName = $fileObject.files[0].full_name

    return $fileName
}


## Rename file with provided id.
function Rename-YammerFilename {

    Param (
        [parameter(Mandatory=$true)][string]$FileId,
        [parameter(Mandatory=$true)][string]$NewFilename
    )
   
    $response = Invoke-WebRequest -Method Put `
                                  -Uri "https://www.yammer.com/api/v1/uploaded_files/$FileId" `
                                  -Headers @{Authorization = "Bearer $authToken"} `
                                  -ContentType "application/x-www-form-urlencoded; charset=UTF-8" `
                                  -Body "name=$NewFilename"

    return $response.StatusCode
}


### Main function ##############################################

while(($authToken -eq $null) -or ($authToken -eq ""))
{
    $authToken = Read-Host -Prompt "Provide the Authorization Token"
}

while(($migrationLogCsv -eq $null) -or ($migrationLogCsv -eq "") -or !([System.IO.File]::Exists($migrationLogCsv)))
{
    $migrationLogCsv = Read-Host -Prompt "Provide the name of the native mode migration error report file (ie: migration_error_report.csv)"
}


Write-Host
Write-Host "Loading report information from $migrationLogCsv ... " -NoNewline
# Loads error report. If error loading the file then stop script execution.
try { $reportLog = Import-Csv $migrationLogCsv } catch { Write-Host $_.Exception.Message -ForegroundColor Red; Break}
Write-Host "Loaded."
Write-Host

# Log file - ScriptName_Results_Timestamp.csv
$logFile = New-Item (($MyInvocation.MyCommand.Name).Split(".")[0] + "_Results_" + (Get-Date -Format FileDateTime) + ".csv") -ItemType File -Force
Add-Content $logFile -Value '"item_id","message","status"'


### messages handling logic START

# Variables for tracking
[string]$errorMsg = ""

# iterate all report entries
foreach($logEntry in $reportLog)
{
    $errorMsg = $logEntry.message

    # the possible scenarios to handle.
    switch -Wildcard ($errorMsg)
    {

        "FileAlreadyExists"
        {
            $currentFilename = Get-YammerFilename -FileId $logEntry.item_id
            $newFilename = "$($logEntry.item_id)_$currentFilename"

            $result = Rename-YammerFilename -FileId $logEntry.item_id -NewFilename $newFilename

            Write-Host "FileAlreadyExists:" $logEntry.item_id -NoNewline

            if($result -eq 200)
            {
                Write-Host " File was renamed from '$currentFilename' to '$newFilename'." -ForegroundColor Green
                Add-Content $logFile -Value """$($logEntry.item_id)"",""FileAlreadyExists"",""File was renamed from '$currentFilename' to '$newFilename'."""
            }
            else ## web exceptions to be catch
            {
                Write-Host " There was an error while trying to rename file '$currentFilename' with Id $($logEntry.item_id). Error: $result" -ForegroundColor Red
                Add-Content $logFile -Value """$($logEntry.item_id)"",""FileAlreadyExists"",""There was an error while trying to rename file '$currentFilename' with Id $($logEntry.item_id). Error: $result"""
            }

            break
        }


        default
        {
            break
        }
    }
}

### messages handling logic END

Write-Host
Write-Host "All entries from the Report Log file processed!" -ForegroundColor Green
Write-Host "Execution results were saved to log file: $logFile" -ForegroundColor Green
Pause
Write-Host
