<#
	.SYNOPSIS
		Calls a given SharePoint REST endpoint given proper credentials for the tenant.
#>

# Declare params
Param(
	[Parameter(Mandatory=$True)]
	[String]$Url,
	 
	[Parameter(Mandatory=$False)]
	[Microsoft.PowerShell.Commands.WebRequestMethod]$Method = [Microsoft.PowerShell.Commands.WebRequestMethod]::Get,
	 
	[Parameter(Mandatory=$True)]
	[String]$Username
)

#Add all the right SharePoint DLLs so we can create our account credentials later on
Add-Type -Path "C:\Program Files\Common Files\microsoft shared\Web Server Extensions\16\ISAPI\Microsoft.SharePoint.Client.dll" 
Add-Type -Path "C:\Program Files\Common Files\microsoft shared\Web Server Extensions\16\ISAPI\Microsoft.SharePoint.Client.Runtime.dll"

# Grab the user's password as a secure string
$secPass = Read-Host -Prompt "Enter your Office 365 account password" -AsSecureString 

# Create the SharePoint Online credentials so we can call the APIs
$creds= New-Object Microsoft.SharePoint.Client.SharePointOnlineCredentials($Username, $secPass)

# Once the creds are assembled, create the web request to call the REST APIs
$request = [System.Net.WebRequest]::Create($Url)
$request.Headers.Add("x-forms_based_auth_accepted", "f") 
$request.Accept = "application/json;odata=verbose"
$request.Method=$Method
$request.Credentials = $creds

#execute the response and see what happens
$response = $request.GetResponse()
$stream = $response.GetResponseStream()
$readStream = New-Object System.IO.StreamReader $stream
$responseData=$readStream.ReadToEnd()

#Return the response as JSON
($responseData | ConvertFrom-JSON).d
