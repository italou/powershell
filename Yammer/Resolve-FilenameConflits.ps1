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
    Resolves invalid characters conflits in files name.
   
.DESCRIPTION
    This script will read from the Export Network Data Files.csv file and check
    if the files name can conflict when migrating to native mode due to invalid
    characters in the name. If a conflit is found, the invalid character will be
    replaced with underscore '_'.
 
    It needs an Authorization Token to be able to perform the rename operations
    as well as the Files.csv file from Yammer Network Data Export.

    The Authorization Token can be obtained by registering a Yammer App
    (see https://developer.yammer.com/docs/app-registration) and on the App page
    "Generate a developer token for this application". Provide this token at the
    variable "$authToken" below.

    The Files.csv file can be obtained from going to the Network Admin Settings,
    then Export Network Data. Make sure to disable "Include attachments", it's
    only needed the Files.csv file, not the files themselves.

    WARNING: DO NOT EDIT AND SAVE THE FILE IN EXCEL, DOING SO CAN CHANGE THE 
             CSV ENCODING AND AFFECT SCRIPT FUNCIONALITY.


    In the end, a CSV log file will be provided with the results.
    Example: "<SCRIPTNAME>_Results_<TIMESTAMP>.csv"


.PARAMETER authToken
    Specifies the Authorization Token needed to perform the request.


.PARAMETER filesCsv
    Specifies the Files.csv filename.


.NOTES
    Version : 1.0
    Updated : 2021-01-27

.LINK  
    https://docs.microsoft.com/en-us/yammer/troubleshoot-problems/troubleshoot-native-mode#how-does-yammer-in-native-mode-handle-file-name-conflicts

#>


# Warnings output
Write-Host "ATTENTION: You should use an Authorization Token from a GA user or Yammer App / Developer Token." -ForegroundColor Yellow
Write-Host "           Also be aware that this script was created to be used along with the Files.csv file" -ForegroundColor Yellow
Write-Host "           provided by Yammer Data Export functionality." -ForegroundColor Yellow
Write-Host
Write-Host "Press CTRL + C to cancel. " -NoNewline; Pause
Write-Host

# SET HERE YOUR AUTHORIZATION TOKEN HERE - SEE DESCRIPTION
$authToken = ""

# SET HERE THE Files.csv FILENAME, IE: Files.csv - SEE DESCRIPTION
$filesCsv = ""


# Invalid Characters to replace.
$invalidChars = "[\uFFFD\~\#\%\&\*\{\}\\\:\<\>\?\/\+\|\'\""]"


# Keeps asking for authorization token while empty.
while(($authToken -eq $null) -or ($authToken -eq ""))
{
    $authToken = Read-Host -Prompt "Provide the Authorization Token"
}

# Keeps asking for Files.csv file while empty or file not existing.
while(($filesCsv -eq $null) -or ($filesCsv -eq "") -or !([System.IO.File]::Exists($filesCsv)))
{
    $filesCsv = Read-Host -Prompt "Provide the filename of the Files Data Export file. (ie: Files.csv)"
}

# Reads Files.csv file.
Write-Host
Write-Host "Importing file information from: $filesCsv"
$importedFiles = Import-Csv $filesCsv

# Sorting Files.csv content for possible duplicate filenames.
# Sorting first by group (ascending) then by name (descending).
$sortedFileList = $importedFiles | Sort-Object @{Expression={[float]$_.group_id}; Descending=$false }, @{Expression={[String]$_.name}; Descending=$true }

# Log file - ScriptName_Results_Timestamp.csv
$logFile = New-Item (($MyInvocation.MyCommand.Name).Split(".")[0] + 
                      "_Results_" + 
                      (Get-Date -Format FileDateTime) +
                      ".csv") -ItemType File -Force

# Adding log file headers
Add-Content $logFile -Value '"file_id","status","result","name","group_id","storage_type"'


# In Yammer "legacy" groups you can have files with the same name (as the IDs will be
# different), so checking for this situation and if same name files found in same group
# then the file_id will be appended to the filename.
$lastFilename = ""
$lastFileGroup = ""

# Iterate each file and validates conditions.
foreach($file in $sortedFileList){

    Write-Host "Checking file ID:" $file.file_id "- Name:" $file.name "- Storage:" $file.storage_type -NoNewline

    # Skipping file if not stored in Yammer (AZURE).
    if($file.storage_type -ne "AZURE")
    {
        # logging and output
        Add-Content $logFile -Value """$($file.file_id)"",""Skipped"",""File not stored in Yammer."",""$($file.name)"",""$($file.group_id)"",""$($file.storage_type)"""
        Write-Host " >> Skipped, file not stored in Yammer."
    }
    else
    {
        # Check file name for invalid characters.
        if($file.name -match $invalidChars) # filename is invalid.
        {
            # Resolve invalid characters in file name.
            $newFilename = $file.name -replace , $invalidChars, "_"

            # Check if current file has same name and group id as last iterated file.
            # If so, then appends file id to filename to avoid duplicated filenames.
            if(($newFilename -eq $lastFilename) -and ($file.group_id -eq $lastFileGroup))
            {
                $response = Invoke-WebRequest -Method Put `                                  -Uri "https://www.yammer.com/api/v1/uploaded_files/$($file.file_id)" `                                  -Headers @{Authorization = "Bearer $authToken"} `
                                  -ContentType "application/x-www-form-urlencoded; charset=UTF-8" `
                                  -Body "name=$($newFilename + "_" + $file.file_id)"
            }
            else
            {
                $response = Invoke-WebRequest -Method Put `                                  -Uri "https://www.yammer.com/api/v1/uploaded_files/$($file.file_id)" `                                  -Headers @{Authorization = "Bearer $authToken"} `
                                  -ContentType "application/x-www-form-urlencoded; charset=UTF-8" `
                                  -Body "name=$newFilename"
            }

            # logging and output
            # -Value '"file_id", "status", "result", "name","group_id", "storage_type"'
            Add-Content $logFile -Value """$($file.file_id)"",""Renamed"",""$($newFilename)"",""$($file.name)"",""$($file.group_id)"",""$($file.storage_type)"""
            Write-Host " >> Name conflit found! File renamed to $newFilename" -ForegroundColor Yellow

            # update last iterated file name.
            $lastFilename = $newFilename
        }
        else # filename is valid.
        {
            # logging and output
            # -Value '"file_id", "status", "result", "name","group_id", "storage_type"'
            Add-Content $logFile -Value """$($file.file_id)"",""Skipped"",""No name conflits found."",""$($file.name)"",""$($file.group_id)"",""$($file.storage_type)"""
            Write-Host " >> Skipped, no name conflits found." -ForegroundColor Green
        }

    }
}


Write-Host
Write-Host "All files from $filesCsv processed!" -ForegroundColor Green
Write-Host "Results saved in log file: $logFile" -ForegroundColor Green
Pause
Write-Host