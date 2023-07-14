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
    Resolves orphan file errors for native mode migration by downloading the
    the files and delete them from Viva Engage (Yammer).
   
.DESCRIPTION
    This script will read the native mode error report log and will check for 
    files identified as orphan files. Any found file will then be downloaded
    to the current location and deleted for Viva Engage (Yammer).

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
    Version   : 1.0
    Updated   : 2023-07-14

.LINK  
    https://learn.microsoft.com/en-us/yammer/troubleshoot-problems/troubleshoot-native-mode

#>


# Warnings output
Write-Host
Write-Host "ATTENTION: You should use an Authorization Token from a GA user or Viva Engage (Yammer) App / Developer Token." -ForegroundColor Yellow
Write-Host "           Also be aware that this script was created to be used along with the native mode migration" -ForegroundColor Yellow
Write-Host "           report log CSV file provided by Viva Engage (Yammer)." -ForegroundColor Yellow
Write-Host
Write-Host "Press CTRL + C to cancel. " -NoNewline; Pause
Write-Host

# Authorization token.
[string]$authToken = ""

# Native mode migration report log file.
[string]$migrationLogCsv = ""


### Main function ###
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
Add-Content $logFile -Value '"file_url","file_name","download_result","deleted_result"'


### orphan files handling logic START

# Variables for tracking
[string]$errorType = ""

# iterate all report entries
foreach($entry in $reportLog)
{
    $errorType = $entry.item_type

    # the possible scenarios to handle.
    switch -Wildcard ($errorType)
    {

        "orphaned_files"
        {
            # get file id.
            $orphanFileId = $entry.item_id
            
            # https://www.yammer.com/api/v1/uploaded_files/123456789.json
            $orphanFileURL = "https://www.yammer.com/api/v1/uploaded_files/" + $orphanFileId + ".json"

            # get uploaded file information.
            try 
            { 
                $orphanFile = Invoke-WebRequest -Method Get -Uri $orphanFileURL -Headers @{Authorization = "Bearer $authToken"}

            } 
            catch 
            {
                Write-Host "ERROR: Skipping file with id $orphanFileId. " -ForegroundColor Red -NoNewline
                Write-Host $_.Exception.Message -ForegroundColor Red

                Add-Content $logFile -Value """$($entry.url)"","""",""$($_.Exception.Message)"","""""
                
                Break
            }
            

            # Donwload orphan file.
            $downloadFilename = (ConvertFrom-JSON $orphanFile.Content).files.name
            $downloadFileURL = (ConvertFrom-JSON $orphanFile.Content).files.download_url

            Write-Host "Downloading file $downloadFilename... " -NoNewline

            $session = New-Object Microsoft.PowerShell.Commands.WebRequestSession

            $cookie = New-Object System.Net.Cookie 
            $cookie.Name = "oauth_token"
            $cookie.Value = $authToken
            $cookie.Domain = "yammer.com"
            $session.Cookies.Add($cookie);

            [string]$logEntry = ""

            if(Test-Path $downloadFilename)
            {
                $fileName = $downloadFilename.Split(".")[0]
                $fileExtension = $downloadFilename.Split(".")[1]

                $newFilename = $fileName + "_" + $((Get-Date -Format FileDateTime)) + "." + $fileExtension

                Invoke-WebRequest $downloadFileURL -WebSession $session -Headers @{Authorization = "Bearer $authToken"} -TimeoutSec 900 -OutFile $newFilename
                Write-Host "File already exists, timestamp added: " -ForegroundColor Yellow -NoNewline
                Write-Host $newFilename -ForegroundColor Green

                $logEntry = """$($entry.url)"",""$($downloadFilename)"",""Downloaded as $($newFilename)"""
            }
            else
            {
                Invoke-WebRequest $downloadFileURL -WebSession $session -Headers @{Authorization = "Bearer $authToken"} -TimeoutSec 900 -OutFile $downloadFilename
                Write-Host "Done." -ForegroundColor Green

                $logEntry = """$($entry.url)"",""$($downloadFilename)"",""Downloaded"""
            }


            # Deleting orphan files, https://learn.microsoft.com/en-us/rest/api/yammer/uploaded_filesid
            # DEL https://www.yammer.com/api/v1/uploaded_files/:file_id
            Write-Host "     Deleting file $downloadFilename... " -NoNewline

            try
            {
                $deleteRequest = Invoke-WebRequest -Uri $orphanFileURL -Method Delete -WebSession $session -Headers @{Authorization = "Bearer $authToken"} -TimeoutSec 900
                Add-Content $logFile -Value ($logEntry + ",""Deleted""")

                Write-Host "Done." -ForegroundColor Green
            }
            catch
            {
                Write-Host "ERROR: Skipping file with id $orphanFileId. " -ForegroundColor Red -NoNewline
                Write-Host $_.Exception.Message -ForegroundColor Red

                Add-Content $logFile -Value ($logEntry + ",""$($_.Exception.Message)""")
                
                break
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
