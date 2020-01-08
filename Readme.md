# GetSPODocLibPerms

1. GetSPODocLibPerms.ps1 ... Enumerate all the folders and files  with unique permissions in SPO Document Library (or OD4B).
2. GetSPODobLibSharedWithUsers.ps1 ... Enumerate all the folders/files and the users those files are shared with in SPO Document Library (or OD4B).

Do not use this PowerShell script from Document Library with large number of files, as that might cause throttling.


## Prerequitesite
You need to download SharePoint Online Client Components SDK to run this script.
https://www.microsoft.com/en-us/download/details.aspx?id=42038

You can also acquire the latest SharePoint Online Client SDK by Nuget as well.

1. You need to access the following site. 
https://www.nuget.org/packages/Microsoft.SharePointOnline.CSOM

2. Download the nupkg.
3. Change the file extension to *.zip.
4. Unzip and extract those file.
5. place them in the specified directory from the code. 


## 1. GetSPODocLibPerms.ps1

Enumerate all the folders and files in SPO Document Library with permissions.
Only the folder or files with unique permission scope will are returned.


### How to Run - parameters

-siteurl ... The target SPO site

-listName ... Document Library Title to get the permission list. 

-username ... (OPTIONAL) The target user to upload the file.

-password ... (OPTIONAL) The password of the above user.

-outfile ... (OPTIONAL) The external file (*.tsv) to save the result.

-startPath ... (OPTIONAL) The start/root folder path to get the permissions underneath.

-UseREST ... (OPTIONAL) $true for Speed Oriented option (BETA). Might not be stable.

-ShowAllItems ... (OPTIONAL) Show all permissions. Non-Unique permissions will be exported from cache of the parent role assignments. Default : $false

Example)
.\GetSPODocLibPerms.ps1 -siteUrl https://tenant.shrarepoint.com/sites/docsite -listName Documents


### Reference
The following URL has the original sample code.
https://blogs.technet.microsoft.com/sharepoint_support/2015/04/04/sharepoint-online-1250/


## 2. GetSPODobLibSharedWithUsers.ps1

Enumerate all the folders/files and the shared with users list in SPO Document Library or OneDrive for Business

The shared users include external users as well. If the file is shared with anonymous access, it will additionally return AnonymousView or AnonymousEdit as user. 

NOTE) please make sure to browse to "Site Contents" - "Site usage" in the command bar to see if the "Shared with external users" report works for your requirements.
The UI report would be much more cost effective. 
Please remember that retrieving all the file list and shared with users list could be heavily...


### How to Run - parameters

-siteurl ... The target SPO or OD4B site.

-listName ... Document Library Title to get the permission list. 

-username ... (OPTIONAL) The target user to upload the file.

-password ... (OPTIONAL) The password of the above user.

-outfile ... (OPTIONAL) The external file (*.tsv) to save the result.

Example)
.\GetSPODocLibPerms.ps1 -siteUrl https://tenant-my.shrarepoint.com/personal/user_tenant_onmicrosoft_com -listName Documents


