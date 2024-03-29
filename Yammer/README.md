# Viva Engage (Yammer)

DISCLAIMER:  
The scripts here are made available to you without any express, implied or statutory warranty,
not even the implied warranty of merchantability or fitness for a particular purpose, or the
warranty of title or non-infringement. The entire risk of the use or the results from the use
of this script remains with you. **Please make sure to review all scripts code and comments!**

### [Resolve-NMCannotOpenFile.ps1](Resolve-NMCannotOpenFile.ps1)

Using the native mode migration error report log file, this script will check for files identified
as not being possible to read/open and will delete those so they don't block native mode migration.  
In the end, a CSV log file will be provided with the results.

***Requires:***

- **Authorization Token**: can be obtained by [registering a Viva Engage (Yammer) App](https://learn.microsoft.com/en-us/rest/api/yammer/app-registration)
and on the App page "Generate a developer token for this application".
- **Native mode migration error report log**.

#### Output
```
.\Resolve-NMCannotOpenFile.ps1

ATTENTION: You should use an Authorization Token from a GA user or Viva Engage (Yammer) App / Developer Token.
           Also be aware that this script was created to be used along with the native mode migration
           report log CSV file provided by Viva Engage (Yammer).

Press CTRL + C to cancel. Press Enter to continue...: 

Provide the Authorization Token: XXXXXXX-xxxxxxxxxxxxxxxxxxxx
Provide the name of the native mode migration error report file (ie: migration_error_report.csv): migration_error_report.csv

Loading report information from migration_error_report.csv ... Loaded.

CannotOpenFile: 125172933 File with Id 125172933 was deleted.
CannotOpenFile: 125283904 File with Id 125283904 was deleted.
CannotOpenFile: 110736655 File with Id 110736655 was deleted.
CannotOpenFile: 65315800 File with Id 65315800 was deleted.
CannotOpenFile: 58258798 File with Id 58258798 was deleted.

All entries from the Report Log file processed!
Execution results were saved to log file: Resolve-NMCannotOpenFile_Results_20210226T2144313088.csv
```
<br />

### [Resolve-NMExistingFiles.ps1](Resolve-NMExistingFiles.ps1)

Using the native mode migration error report log file, this script will check for files identified
as already existing in Sharepoint and append the file id on those so every file is migrated to
Sharepoint. In Viva Engage (Yammer) it's possible to have different files with same name, only the file ids will
be different, so this way we can try to migrate all files.  
In the end, a CSV log file will be provided with the results.

***Requires:***

- **Authorization Token**: can be obtained by [registering a Viva Engage (Yammer) App](https://learn.microsoft.com/en-us/rest/api/yammer/app-registration)
and on the App page "Generate a developer token for this application".
- **Native mode migration error report log**: ie: migration_error_report.csv.

#### Output
```
.\Resolve-NMExistingFiles.ps1

ATTENTION: You should use an Authorization Token from a GA user or Viva Engage (Yammer) App / Developer Token.
           Also be aware that this script was created to be used along with the native mode migration
           report log CSV file provided by Viva Engage (Yammer).

Press CTRL + C to cancel. Press Enter to continue...: 

Provide the Authorization Token: XXXXXXX-xxxxxxxxxxxxxxxxxxxx
Provide the name of the native mode migration error report file (ie: migration_error_report.csv): migration_error_report.csv

Loading report information from migration_error_report.csv ... Loaded.

FileAlreadyExists: 137802295 File was renamed from 'DuplicatedFile' to '137802295_DuplicatedFile'.
FileAlreadyExists: 136811551 File was renamed from 'DuplicatedFile' to '136811551_DuplicatedFile'.
FileAlreadyExists: 65315927 File was renamed from 'DuplicatedFile' to '65315927_DuplicatedFile'.

All entries from the Report Log file processed!
Execution results were saved to log file: Resolve-NMExistingFiles_Results_20210226T2148293670.csv
```
<br />

### [Resolve-NMFilenameConflits.ps1](Resolve-NMFilenameConflits.ps1)

This script will read from the "Export Network Data" Files.csv file and check if the files name can
conflict when migrating to native mode due to invalid characters in the name. If a conflit is
found, the invalid character will be replaced with underscore '_'. If the new filename ends up being
a duplicated, then the file id will be also appended to filename.  
In the end, a CSV log file will be provided with the results.

***Requires:***

- **Authorization Token**: can be obtained by [registering a Viva Engage (Yammer) App](https://learn.microsoft.com/en-us/rest/api/yammer/app-registration)
and on the App page "Generate a developer token for this application".
- **Files.csv**: Can be obtained from going to the Network Admin Settings, then Export Network Data.
Make sure to disable "Include attachments" as it's only needed the information about the files.

#### Output
```
.\Resolve-NMFilenameConflits.ps1

ATTENTION: You should use an Authorization Token from a GA user or Viva Engage (Yammer) App / Developer Token.
           Also be aware that this script was created to be used along with the Files.csv file
           provided by Viva Engage (Yammer) Data Export functionality.

Press CTRL + C to cancel. Press Enter to continue...: 

Provide the Authorization Token: XXXXXXX-xxxxxxxxxxxxxxxxxxxx
Provide the filename of the Files Data Export file. (ie: Files.csv): Files.csv

Importing file information from: Files.csv
Checking file ID: 889069027328 - Name: YammerDoc1.docx - Storage: AZURE >> Skipped, no name conflits found.
Checking file ID: 889069101056 - Name: Sp3c!alÇÇÂõ+�ller.docx - Storage: AZURE >> Name conflit found! File renamed to Sp3c!alÇÇÂõ__ller.docx
Checking file ID: 889068978176 - Name: Sp3c!alÇÇÂõ+�ller.docx - Storage: AZURE >> Name conflit found! File renamed to 889068978176_Sp3c!alÇÇÂõ__ller.docx
Checking file ID: 889068830720 - Name: Sp3c!alÇÇÂõ+0ller.docx - Storage: SHAREPOINT >> Skipped, file not stored in Yammer.

All files from Files.csv processed!
Results saved in log file: Resolve-NMFilenameConflits_Results_20210226T2237440310.csv
```
<br />

### [Resolve-NMOrphanFiles.ps1](Resolve-NMOrphanFiles.ps1)

This script will read the native mode error report log and will check for files identified as 
orphan files. Any found file will then be downloaded to the current location and deleted for 
Viva Engage (Yammer).
In the end, a CSV log file will be provided with the results.

***Requires:***

- **Authorization Token**: can be obtained by [registering a Viva Engage (Yammer) App](https://learn.microsoft.com/en-us/rest/api/yammer/app-registration)
and on the App page "Generate a developer token for this application".
- **Native mode migration error report log**: ie: migration_error_report.csv.

#### Output
```
.\Resolve-NMOrphanFiles.ps1

ATTENTION: You should use an Authorization Token from a GA user or Viva Engage (Yammer) App / Developer Token.
           Also be aware that this script was created to be used along with the Files.csv file
           provided by Yammer Data Export functionality.

Press CTRL + C to cancel. Press Enter to continue...: 

Provide the Authorization Token: XXXXXXX-xxxxxxxxxxxxxxxxxxxx
Provide the name of the native mode migration error report file (ie: migration_error_report.csv): migration_error_report.csv

Loading report information from migration_error_report.csv ... Loaded.

Downloading file Dummy PDF 2.PDF... Done.
     Deleting file Dummy PDF 2.PDF... Done.
Downloading file pre-travel_checklist.one... File already exists, timestamp added: pre-travel_checklist_20230714T1231038626.one
     Deleting file pre-travel_checklist.one... Done.
ERROR: Skipping file with id 1753581019163. The remote server returned an error: (404) Not Found.
Downloading file Dummy Photo 2.JPG... File already exists, timestamp added: Dummy Photo 2_20230714T1231052972.JPG
     Deleting file Dummy Photo 2.JPGe... Done.
Downloading file Dummy Photo 4.jpg... Done.
     Deleting file Dummy Photo 4.jpg... Done.
Downloading file tf02930020.xltx... Done.
     Deleting file tf02930020.xltx... Done.

All entries from the Report Log file processed!
Execution results were saved to log file: Resolve-NMOrphanFiles_Results_20230714T1231010633.csv
```

