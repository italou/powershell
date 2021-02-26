# SharepointOnline

DISCLAIMER:  
The scripts here are made available to you without any express, implied or statutory warranty,
not even the implied warranty of merchantability or fitness for a particular purpose, or the
warranty of title or non-infringement. The entire risk of the use or the results from the use
of this script remains with you. **Please make sure to review all scripts code and comments!**

### [Clear-ListItemsRetentionLabel.ps1](SharepointOnline/Clear-ListItemsRetentionLabel.ps1)

This script will iterate through all of the items in the specified library and will clear the
retention label applied to the items (compliance tag). It retrieves list items with pagination,
set to 2500 items per page in order to avoid list view threshold limit (5K items) and will
execute the update query in batches of 150 items to avoid "large message" error.
It can be also used with ODB personal sites, site collection url will be the user personal site
url (ie: personal/user_contoso_com) and the library name will the default Documents library.

***Requires: Sharepoint Online Client Components***

<br />

### [Clear-ListItemsRetentionLabelPnP.ps1](SharepointOnline/Clear-ListItemsRetentionLabelPnP.ps1)

Does the same as above but it's based on PnP. This one supports MFA but it's a bit slower due
to read every item in the list, one by one. Also uses pagination, set to 2500 items.

***Requires: PnP.PowerShell module***

