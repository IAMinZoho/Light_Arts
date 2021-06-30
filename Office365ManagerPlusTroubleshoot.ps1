$MyUsernameDomain=Read-Host "Please enter Office 365 UserName(Global Admin)"
$SecurePassword=Read-Host "Please Enter Password" -AsSecureString

$fullMessage="";
if (get-service -Name msoidsvc) 
{
    $fullMessage=$fullMessage+"`n"+"MSOnline Sign Assistant - Available";
}
else
{

    $fullMessage=$fullMessage+"`n"+"MSOnline Sign Assistant - Not Available";
}
 
if([Activator]::CreateInstance([Type]::GetTypeFromCLSID([Guid]'{DCB00C01-570F-4A9B-8D69-199FDBA5723B}')).IsConnectedToInternet)
{
     $fullMessage=$fullMessage+"`n"+"Internet Connection - Available";
}else { 
    $fullMessage=$fullMessage+"`n"+"Internet Connection - Not available";
 }

$version="";
$ndpDirectory = 'hklm:\SOFTWARE\Microsoft\NET Framework Setup\NDP\'
$v4Directory = "$ndpDirectory\v4\Full"
if (Test-Path $v4Directory) 
{
    $version = Get-ItemProperty $v4Directory -name Version | select -expand Version
    Write-host ".Net Highest version available-" $version
}
 
 
$dotNetStatus=$PSVersiontable.CLRVersion;
Write-Host "CLRVersion -"$dotNetStatus;

if($dotNetStatus.Major -ge 4)
{
     $fullMessage=$fullMessage+"`n"+"Required .Net v4+ - Available";         
}
else
{
     $fullMessage=$fullMessage+"`n"+"Required .Net v4+ - Not Available";
     $fullMessage=$fullMessage+"`n"+"Available .Net version -"+$dotNetStatus;
}

$PSVersionNumber=$PSVersiontable.PSversion.Major;
Write-Host "PS version -"$PSVersiontable.PSversion;

if($PSVersionNumber -ge 3)
{
     $fullMessage=$fullMessage+"`n"+"Required PS V3.0 - Available";
}
else
{
     $fullMessage=$fullMessage+"`n"+"Required PS V3.0 - Not Available";
     $fullMessage=$fullMessage+"`n"+"Available PS version -"+$PSVersiontable.PSversion;
}

$SystemDetails=Get-WmiObject Win32_OperatingSystem|select Caption,CSDVersion,OSArchitecture

Write-Host "Operating System - " $SystemDetails.Caption;

if($SystemDetails.OSArchitecture -Match "64")
{
	$fullMessage=$fullMessage+"`n"+"Required OS Architecture - Available";
}
else
{	
	$fullMessage=$fullMessage+"`n"+"Required OS Architecture - Not Available";
}

Write-Host "OS Architecture  - " $SystemDetails.OSArchitecture;

if($SystemDetails.CSDVersion -ne $null)
{
	Write-Host "Service Pack - "$SystemDetails.CSDVersion ;
}
$powershell64BitVersion=[Environment]::Is64BitProcess;
if($powershell64BitVersion)
{
	$fullMessage=$fullMessage+"`n"+"Required Powershell 64-Bit - Available";
}
else
{	
	$fullMessage=$fullMessage+"`n"+"Required Powershell 64-Bit - Not Available";
}
Write-Host "Is 64-Bit Powershell - " $powershell64BitVersion;

$culture = Get-Culture;
if($culture -ne 1033)
{
	Write-Host "Powershell Culture LCID- " $culture.LCID;
}

if (Get-Module -ListAvailable -Name msonline) 
{
    $msolModules = (Get-Module -ListAvailable -Name MSOnline)
    $fullMessage=$fullMessage+"`n"+"MSOnline module- Available. Version : "+($msolModules.version -Join ',');
    $hasModule=$true;
}
else
{
    $fullMessage=$fullMessage+"`n"+"MSOnline module- Not Available";
    $hasModule=$false;
}
if (Get-Module -ListAvailable -Name AzureAD) 
{
    $fullMessage=$fullMessage+"`n"+"AzureAD v2 module- Available. Version : "+((Get-Module -ListAvailable -Name AzureAD).Version).ToString();
    $hasModule=$true;
}
else
{
    $fullMessage=$fullMessage+"`n"+"AzureAD v2 module- Not Available";
    $hasModule=$false;
}

if (Get-Module -ListAvailable -Name SkypeOnlineConnector) 
{
    $fullMessage=$fullMessage+"`n"+"SkypeOnlineConnector module- Available. Version : "+((Get-Module -ListAvailable -Name SkypeOnlineConnector).Version).ToString();
    $hasSkypeModule=$true;
}
else
{
    $fullMessage=$fullMessage+"`n"+"SkypeOnlineConnector module- Not Available";
    $hasSkypeModule=$false;
}

$credential =New-object System.Management.Automation.PSCredential $MyUsernameDomain,$SecurePassword

Write-host "Creating new session...";
$Session = New-PSSession  -ConfigurationName Microsoft.Exchange -ConnectionUri https://outlook.office365.com/powershell-liveid/ -Credential $credential -Authentication Basic -AllowRedirection;

Write-host "Checking execution policy...";
Set-ExecutionPolicy -ExecutionPolicy RemoteSigned -Force -Scope process;

$sessionStatus=$False;

Try {

     Write-host "Creating PSsession...";
     Import-PSSession -Session $Session -AllowClobber;
     $fullMessage=$fullMessage+"`n"+"Exchange Session - Success";
     
     }
     Catch {
         
          $fullMessage=$fullMessage+"`n"+"Exchange Session - Error Occurred";
	}

	
if( $hasModule)
{
try
{
Import-module MSOnline;
    Connect-msolservice -Credential $credential -EA stop;
    $sessionStatus=$True;
     $fullMessage=$fullMessage+"`n"+"MSOnline Session - Success";
}
catch
{
 $sessionStatus=$False;
   $fullMessage=$fullMessage+"`n"+"MSOnline Session - Error Occurred";

}
}
else
{
$fullMessage=$fullMessage+"`n"+"MSOnline Session -Unable to verify. Windows Azure module must be installed.";
}

if( $hasSkypeModule)
{
try
{
	Import-module SkypeOnlineConnector;
	$session = New-CsOnlineSession  -Credential $credential;
	Import-PSSession $session -AllowClobber;
     $fullMessage=$fullMessage+"`n"+"SkypeOnline Session - Success";
}
catch
{
   $fullMessage=$fullMessage+"`n"+"SkypeOnline Session - Error Occurred";

}
}
else
{
$fullMessage=$fullMessage+"`n"+"SkypeOnline Session -Unable to verify. SkypeOnlineConnector must be installed.";
}
if( $hasModule -and $sessionStatus)
{

	
	$role = Get-MsolRole -RoleName “Company Administrator”;
	$EmailAdd=@();
	$EmailAdd=Get-MsolRoleMember -RoleObjectId $role.ObjectId;
	
	if($EmailAdd.EmailAddress -contains $MyUsernameDomain)
	{
	$fullMessage=$fullMessage+"`n"+"Is Global Admin Account - True";
	}
    else
    {
        $fullMessage=$fullMessage+"`n"+"Is Global Admin Account - False";
    }
}elseif($hasModule -and -not $sessionStatus)
{
$fullMessage=$fullMessage+"`n"+"Is Global Admin Account - Unable to verify. Could not connect to Azure.";
}
elseif(-not $hasModule)
{
$fullMessage=$fullMessage+"`n"+"Is Global Admin Account - Unable to verify. Windows Azure module must be installed.";
}


Write-Host "`n================================================="
Write-Host "`nOffice 365 Manager Plus Requirement Status"
Write-Host "`n================================================="

$fullMessage

