<#
 This Sample Code is provided for the purpose of illustration only and is not intended to be used in a production environment. 
 THIS SAMPLE CODE AND ANY RELATED INFORMATION ARE PROVIDED "AS IS" WITHOUT WARRANTY OF ANY KIND, EITHER EXPRESSED OR IMPLIED, 
 INCLUDING BUT NOT LIMITED TO THE IMPLIED WARRANTIES OF MERCHANTABILITY AND/OR FITNESS FOR A PARTICULAR PURPOSE.  

 We grant you a nonexclusive, royalty-free right to use and modify the sample code and to reproduce and distribute the object 
 code form of the Sample Code, provided that you agree: 
    (i)   to not use our name, logo, or trademarks to market your software product in which the sample code is embedded; 
    (ii)  to include a valid copyright notice on your software product in which the sample code is embedded; and 
    (iii) to indemnify, hold harmless, and defend us and our suppliers from and against any claims or lawsuits, including 
          attorneys' fees, that arise or result from the use or distribution of the sample code.

Please note: None of the conditions outlined in the disclaimer above will supercede the terms and conditions contained within 
             the Premier Customer Services Description.
#>
param(
  [Parameter(Mandatory=$true)]
  [string]$siteUrl,
  [Parameter(Mandatory=$true)]
  [string]$listName,
  [Parameter(Mandatory=$false)]
  $startPath,
  [Parameter(Mandatory=$false)]
  $username,
  [Parameter(Mandatory=$false)]
  $password,
  [Parameter(Mandatory=$false)]
  $ShowAllItems=$false,
  [Parameter(Mandatory=$false)]
  $UseREST=$false,
  [Parameter(Mandatory=$false)]
  $outfile
)

Add-Type -Path "C:\Program Files\Common Files\microsoft shared\Web Server Extensions\16\ISAPI\Microsoft.SharePoint.Client.dll"
Add-Type -Path "C:\Program Files\Common Files\microsoft shared\Web Server Extensions\16\ISAPI\Microsoft.SharePoint.Client.Runtime.dll"

$UserAgentString = "NONISV|Contoso|Application/1.0"

$script:context = new-object Microsoft.SharePoint.Client.ClientContext($siteUrl)
$pwd = $null
if ($username -eq $null)
{
  $cred = Get-Credential
  $username = $cred.UserName
  $pwd = $cred.Password
}
else
{
  $pwd = convertto-securestring $password -AsPlainText -Force
}
$credentials = New-Object Microsoft.SharePoint.Client.SharePointOnlineCredentials($username, $pwd)
$script:context.Credentials = $credentials

$cookie = $null
if ($UseREST)
{
  $rootUrl = ("https://" + (New-Object system.uri($siteUrl)).Host);
  $cookie = $context.Credentials.GetAuthenticationCookie($rootUrl, $true);
  $script:UseREST = $true
  $script:restsession = New-Object Microsoft.PowerShell.Commands.WebRequestSession;
  $script:restsession.Cookies.SetCookies($rootUrl, $cookie);
  $script:restheaders = @{
    "UserAgent" = $UserAgentString
  }

}

$script:context.add_ExecutingWebRequest({
  param ($source, $eventArgs);
  $request = $eventArgs.WebRequestExecutor.WebRequest;
  $request.UserAgent = $UserAgentString;
})

function ExecuteQueryWithIncrementalRetry {
  param (
      [parameter(Mandatory = $true)]
      [int]$retryCount
  );

  $DefaultRetryAfterInMs = 120000;
  $RetryAfterHeaderName = "Retry-After";
  $retryAttempts = 0;

  if ($retryCount -le 0) {
      throw "Provide a retry count greater than zero."
  }

  while ($retryAttempts -lt $retryCount) {
      try {
          $script:context.ExecuteQuery();
          return;
      }
      catch [System.Net.WebException] {
          $response = $_.Exception.Response

          if (($null -ne $response) -and (($response.StatusCode -eq 429) -or ($response.StatusCode -eq 503))) {
              $retryAfterHeader = $response.GetResponseHeader($RetryAfterHeaderName);
              $retryAfterInMs = $DefaultRetryAfterInMs;

              if (-not [string]::IsNullOrEmpty($retryAfterHeader)) {
                  if (-not [int]::TryParse($retryAfterHeader, [ref]$retryAfterInMs)) {
                      $retryAfterInMs = $DefaultRetryAfterInMs;
                  }
                  else {
                      $retryAfterInMs *= 1000;
                  }
              }

              Write-Output ("CSOM request exceeded usage limits. Sleeping for {0} seconds before retrying." -F ($retryAfterInMs / 1000))
              #Add delay.
              Start-Sleep -m $retryAfterInMs
              #Add to retry count.
              $retryAttempts++;
          }
          else {
              throw;
          }
      }
  }

  throw "Maximum retry attempts {0}, have been attempted." -F $retryCount;
}


###
# ref : https://sharepoint.stackexchange.com/questions/126221/spo-retrieve-hasuniqueroleassignements-property-using-powershell
##
Function Invoke-LoadMethod() {
  param(
     [Microsoft.SharePoint.Client.ClientObject]$Object = $(throw "Please provide a Client Object"),
     [string]$PropertyName
  ) 
     $ctx = $Object.Context
     $load = [Microsoft.SharePoint.Client.ClientContext].GetMethod("Load") 
     $type = $Object.GetType()
     $clientLoad = $load.MakeGenericMethod($type) 
  
  
     $Parameter = [System.Linq.Expressions.Expression]::Parameter(($type), $type.Name)
     $Expression = [System.Linq.Expressions.Expression]::Lambda(
              [System.Linq.Expressions.Expression]::Convert(
                  [System.Linq.Expressions.Expression]::PropertyOrField($Parameter,$PropertyName),
                  [System.Object]
              ),
              $($Parameter)
     )
     $ExpressionArray = [System.Array]::CreateInstance($Expression.GetType(), 1)
     $ExpressionArray.SetValue($Expression, 0)
     $clientLoad.Invoke($ctx,@($Object,$ExpressionArray))
  }

  function WriteOut($text, $append)
  {
    if ($outfile -eq $null)
    {
      Write-Output $text
    }
    else
    {
      if ($append)
      {
        $text | Out-File $outfile -Append -Encoding UTF8
      }
      else
      {
        $text | Out-File $outfile -Encoding UTF8
      }
    }
  }

function GetObjectRoles($obj, $url)
{
  $script:context.Load($obj.RoleAssignments)
  ExecuteQueryWithIncrementalRetry -retryCount 5 
  $aclList = New-Object System.Collections.ArrayList

  foreach ($ra in $obj.RoleAssignments)
  {
    $ace = New-Object PSObject
    $script:context.Load($ra.Member)
    ExecuteQueryWithIncrementalRetry -retryCount 5 

    $cnt = Add-Member -InputObject $ace -MemberType NoteProperty -Name UserName -Value $ra.Member.Title

    $script:context.Load($ra.RoleDefinitionBindings)
    ExecuteQueryWithIncrementalRetry -retryCount 5 

    $i = 0
    $defs = New-Object System.Text.StringBuilder
    foreach ($rd in $ra.RoleDefinitionBindings)
    {
      if ($i -ne 0)
      {
        $cnt = $defs.Append(",")
      }
      $cnt = $defs.Append($rd.Name)
      $i++
    }
    $cnt = Add-Member -InputObject $ace -MemberType NoteProperty -Name RoleDefinitions -Value $defs.ToString()
    $cnt = $aclList.Add($ace)
  }
  return $aclList
}

function GetObjectRolesWithREST
{
  param(
    [parameter(Mandatory=$true)]
    [ValidateSet("List","Folder","File")]
    $objtype, 
    $url
  )

  switch ($objtype)
  {
    "List" {
      $restEndPoint = $siteUrl.TrimEnd("/") + "/_api/web/getlist(@path)/roleassignments?`$expand=Member,RoleDefinitionBindings&@path='" + $url + "'"
      break;
    }
    "Folder" {
      $restEndPoint = $siteUrl.TrimEnd("/") + "/_api/web/getfolderbyserverrelativeurl(@path)/listitemallfields/roleassignments?`$expand=Member,RoleDefinitionBindings&@path='" + $url + "'"
      break;
    }
    "File" {
      $restEndPoint = $siteUrl.TrimEnd("/") + "/_api/web/getfilebyserverrelativeurl(@path)/listitemallfields/roleassignments?`$expand=Member,RoleDefinitionBindings&@path='" + $url + "'"
      break;
    }
  }

  $restresult = Invoke-RestMethod -Method Get -Uri $restEndPoint -Headers $script:restheaders -WebSession $script:restsession;
  $aclList = New-Object System.Collections.ArrayList

  foreach ($entry in $restresult)
  {
    $ace = New-Object PSObject
    $cnt = Add-Member -InputObject $ace -MemberType NoteProperty -Name UserName -Value $entry.link[1].inline.entry.content.properties.Title

    $i = 0
    $defs = New-Object System.Text.StringBuilder
    foreach ($rd in $entry.link[2].inline.feed.entry.content.properties.Name)
    {
      if ($i -ne 0)
      {
        $cnt = $defs.Append(",")
      }
      $cnt = $defs.Append($rd)
      $i++
    }
    $cnt = Add-Member -InputObject $ace -MemberType NoteProperty -Name RoleDefinitions -Value $defs.ToString()
    $cnt = $aclList.Add($ace)
  }
  return $aclList
}


function PrintObjectRoles($aclList, $url)
{
  foreach ($ace in $aclList)
  {
    $sb = new-object System.Text.StringBuilder
    $cnt = $sb.Append($url)
    $cnt = $sb.Append("`t")

    $cnt = $sb.Append($ace.UserName)
    $cnt = $sb.Append("`t")
    $cnt = $sb.Append($ace.RoleDefinitions)
    WriteOut -text $sb.ToString() -append $true
  }
}


function EnumPermsInFolder($list, $ServerRelativeUrl, $parentAcL)
{
  do
  {
    $camlQuery = New-Object Microsoft.SharePoint.Client.CamlQuery
    $camlQuery.ListItemCollectionPosition = $position
    $camlQuery.ViewXml = "<View><RowLimit>5000</RowLimit></View>";
    if ($serverRelativeUrl -ne $null)
    {
       $camlQuery.FolderServerRelativeUrl = $ServerRelativeUrl
    }
    $listItems = $list.GetItems($camlQuery);
    $script:context.Load($listItems);
    ExecuteQueryWithIncrementalRetry -retryCount 5 

    foreach($listItem in $listItems)
    {

      Invoke-LoadMethod -Object $listItem -PropertyName "HasUniqueRoleAssignments"
      ExecuteQueryWithIncrementalRetry -retryCount 5 
    
      if ($listItem.FileSystemObjectType -eq [Microsoft.SharePoint.Client.FileSystemObjectType]::Folder)
      {
        $script:context.Load($listItem.Folder)
        ExecuteQueryWithIncrementalRetry -retryCount 5 
        if ($listItem.HasUniqueRoleAssignments)
        {
          if ($UseREST)
          {
            $aclList = GetObjectRolesWithREST -objtype "Folder" -url $listItem.Folder.ServerRelativeUrl
          }
          else {
            $aclList = GetObjectRoles -obj $listItem -url $listItem.Folder.ServerRelativeUrl
          }
          PrintObjectRoles -aclList $aclList -url $listItem.Folder.ServerRelativeUrl
          EnumPermsInFolder -List $list -ServerRelativeUrl $listItem.Folder.ServerRelativeUrl -parentACL $aclList
        }
        else
        {
          if ($ShowAllItems)
          {
            PrintObjectRoles -aclList $parentACL -url $listItem.Folder.ServerRelativeUrl
          }
          EnumPermsInFolder -List $list -ServerRelativeUrl $listItem.Folder.ServerRelativeUrl -parentACL $parentAcl
        }
      }
      else {
        if ($listItem.HasUniqueRoleAssignments)
        {
          $script:context.Load($listItem.File)
          ExecuteQueryWithIncrementalRetry -retryCount 5 
          if ($useREST)
          {
            $aclList = GetObjectRolesWithREST -objtype "File" -url $listItem.File.ServerRelativeUrl
          }
          else {
            $aclList = GetObjectRoles -obj $listItem -url $listItem.File.ServerRelativeUrl
          }
          PrintObjectRoles -aclList $aclList -url $listItem.File.ServerRelativeUrl
        }
        else
        {
          if ($ShowAllItems)
          {
            $script:context.Load($listItem.File)
            ExecuteQueryWithIncrementalRetry -retryCount 5 
            PrintObjectRoles -aclList $parentACL -url $listItem.File.ServerRelativeUrl
          }
        }
      }
    } 
    $position = $listItems.ListItemCollectionPosition
  }
  while($position -ne $null)
}


$list = $script:context.Web.Lists.GetByTitle($listName)
$script:context.Load($list)
$script:context.Load($list.RootFolder)
ExecuteQueryWithIncrementalRetry -retryCount 5 

WriteOut -text "Scope`tRoleAssignment`tRoleDefinition" 
if ($startPath -eq $null)
{
  $startPath = $list.RootFolder.ServerRelativeUrl
}

if ($UseREST)
{
  $aclList = GetObjectRolesWithREST -objtype "List" -url $startPath
}
else {
  $aclList = GetObjectRoles -obj $list -url $startPath
}
PrintObjectRoles -aclList $aclList -url $startPath
EnumPermsInFolder -List $list -serverRelativeUrl $startPath -parentACL $aclList

