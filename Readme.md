# GetSPODocLibPerms

Enumerate all the folders and files in SPO Document Library with permissions on them.
Only the folder or files which has unique permission scope will be extracted.

Do not use this PowerShell script to get permission from large number of files, as that might cause throttling.


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

## How to Run - parameters

-siteurl ... The target SPO site

-listTitle ... Document Library Title to get the permission list. 

-username ... (OPTIONAL) The target user to upload the file.

-password ... (OPTIONAL) The password of the above user.

-outfile ... (OPTIONAL) The external file (*.tsv) to save the result.


## Reference
The following URL has the original sample code.
https://blogs.technet.microsoft.com/sharepoint_support/2015/04/04/sharepoint-online-1250/
