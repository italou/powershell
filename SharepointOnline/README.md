# SharepointOnline

DISCLAIMER:  
The scripts here are made available to you without any express, implied or statutory warranty, not even the implied warranty of merchantability or fitness for a particular purpose, or the warranty of title or non-infringement. The entire risk of the use or the results from the use of this script remains with you. **Please make sure to review all scripts code and comments!**

### [Clear-ListItemsRetentionLabel.ps1](Clear-ListItemsRetentionLabel.ps1)

This script will iterate through all of the items in the specified library and will clear the retention label applied to the items (compliance tag). It retrieves list items with pagination, set to 2500 items per page in order to avoid list view threshold limit (5K items) and will execute the update query in batches of 150 items to avoid "large message" error.
It can be also used with ODB personal sites, site collection url will be the user personal site url (ie: personal/user_contoso_com) and the library name will the default Documents library.

***Requires: Sharepoint Online Client Components***

#### Output
```
.\Clear-ListItemsRetentionLabel.ps1
Enter username: user@contoso.onmicrosoft.com
Enter site collection URL: https://contoso-my.sharepoint.com/personal/user_contoso_onmicrosoft_com
Enter list/library name: Documents

Getting list... Done!

Getting list items... Done!
'Folder1' queued to clear label 'Block Deletion Forever'.
'Document1.docx' queued to clear label 'Block Deletion Forever'.
'Document2.docx' queued to clear label 'Block Deletion Forever'.
'Document.docx' queued to clear label 'Block Deletion Forever'.
'Document.docx' queued to clear label 'Block Deletion Forever'.

Updating queued items... Done

A total of 5 items were cleared in 'Documents'.
```
<br />

### [Clear-ListItemsRetentionLabelPnP.ps1](Clear-ListItemsRetentionLabelPnP.ps1)

Does the same as the one above but it uses PnP. Supports MFA but it's a bit slower due to scan every item in the list, one by one. It also uses pagination, set to 2500 items.

***Requires: PnP.PowerShell module***

#### Output
```
.\Clear-ListItemsRetentionLabelPnP.ps1
Enter site collection URL: https://contoso-my.sharepoint.com/personal/user_contoso_onmicrosoft_com
Enter list/library name: Documents

Getting list items... Done!
Clearing label 'Block Deletion Forever' on item 'Folder1'... Cleared!
Clearing label 'Block Deletion Forever' on item 'Document2.docx'... Cleared!
Clearing label 'Block Deletion Forever' on item 'Document.docx'... Cleared!
Clearing label 'Block Deletion Forever' on item 'Document.docx'... Cleared!

A total of 4 items were cleared in Documents.
```
<br />

### [Enforce-MajorVersionLimit.ps1](Enforce-MajorVersionLimit.ps1)

This script iterates through all the libray items and for files it checks it the file has more versions than the set major versions limit.
If so, it removes the oldest versions until it has the same number of versions as the set limit.

***Requires: PnP.PowerShell module***

#### Output
```
.\Enforce-MajorVersionLimit.ps1
Enter site collection URL: https://contoso.sharepoint.com/sites/mySite
Enter library name: Documents
Documents has limit of 100 major versions.
Getting library items...
Processing library items...
Checking file /sites/mySite/Shared Documents/Versions_10.txt... 9 versions found.
Checking file /sites/mySite/Shared Documents/Versions_110.txt... 109 versions found!
   Removing 9 versions. File now only has 100 versions.
Checking file /sites/mySite/Shared Documents/Versions_100.txt... 99 versions found.
Checking file /sites/mySite/Shared Documents/Versions_105.txt... 104 versions found!
   Removing 4 versions. File now only has 100 versions.
Checking file /sites/mySite/Shared Documents/Versions_150.txt... 149 versions found!
   Removing 49 versions. File now only has 100 versions.
Checking file /sites/mySite/Shared Documents/Versions_123.txt... 122 versions found!
   Removing 22 versions. File now only has 100 versions.
```
