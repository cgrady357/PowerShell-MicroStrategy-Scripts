<#
        .NAME
		 Check-AvailabilityMsiWeb
		.DESCRIPTION
         Checks that a user can successfully log in MicroStrategy Web.  It will send an email alert if script can't log into web.
        .PARAMETER UrlToCheck
	     The MicroStrategy Web Url checked for availability
		.PARAMETER UserID
		 The user id used to login into MicroStrategy
        .PARAMETER Password
		 The password used to login into MicroStrategy
        .PARAMETER WebAppServer
		 The server that the web application server runs on
		.PARAMETER IntelligenceServer
		 The server that the MicroStrategy Intelligence Server runs on
		.PARAMETER ProjectToVerifyIsAvailable
		 The project used to test that MicroStrategy Web is available (i.e. Enterprise Manager)
		.PARAMETER SmtpServer 
		 The SmtpServer used to send out email after script completion
        .PARAMETER From 
		 The From email address used to send out an email after script completion
		.PARAMETER To
		 The To email address used to send out an email after script completion
	    .PARAMETER ControlMJobName
		 The control m job name that calls this script
		.PARAMETER JobFunction
		 Explains what this script does
		.PARAMETER ScriptLog
         The log file for this script		

        .NOTES
         AUTHOR....:  Craig Grady
         LAST EDIT.:  09/09/2014
         CREATED...:  08/14/2014
#>

[CmdletBinding()]
param(
[Parameter(Mandatory=$true)]
[string]$UrlToCheck,
[Parameter(Mandatory=$true)]
[string]$UserID,
[Parameter(Mandatory=$true)]
[string]$Password,
[Parameter(Mandatory=$true)]
[string]$WebAppServer,
[Parameter(Mandatory=$true)]
[string]$IntelligenceServer,
[Parameter(Mandatory=$true)]
[string]$ProjectToVerifyIsAvailable,
[Parameter(Mandatory=$true)]
[string]$SmtpServer,
[Parameter(Mandatory=$true)]
[string]$From,
[Parameter(Mandatory=$true)]
[string]$To, 
[Parameter(Mandatory=$true)]
[string]$ControlMJobName,
[string]$JobFunction = "Checks that MicroStrategy Web and the MicroStrategy Intelligence Server are accessible",
[Parameter(Mandatory=$true)]
[string]$ScriptLogPath,
[string]$ScriptLog = $ControlMJobName + $(get-date -f "-yyyy-MM-dd") + ".log" 
)

if( Test-Path -Path $ScriptLogPath) { 
    $ScriptLog = (Join-Path -Path $ScriptLogPath -ChildPath $ScriptLog)
	} 
	else { 
	$ScriptLog = (Join-Path -Path $Env:tmp   -ChildPath $ScriptLog) 
    }

$ScriptStartTime = Get-Date
$nl = [Environment]::NewLine
$Alert = $false
$Status = $null
$LocalHost = $env:ComputerName

$url = $UrlToCheck
$r=Invoke-WebRequest $url -SessionVariable msi
$form = $r.Forms[0]
$form.fields["Uid"]=$UserID
$form.fields["Pwd"]=$Password
$r=Invoke-WebRequest -Uri ($url + ";" + $form.Action.Split(";")[1]) -WebSession $msi -Method POST -Body $form.Fields
if($r.links.InnerText -contains $ProjectToVerifyIsAvailable) {$Status = "Available" } else {$Status = "Not Available"; $Alert = $true}

$LogMessage = @"
Script Start Time: $($ScriptStartTime) 
Url Checked: $($UrlToCheck)
MicroStrategy Web Availability Status: $($Status)
MicroStrategy Web Application Server: $($WebAppServer)
MicroStrategy Intelligence Server: $($IntelligenceServer)
Server Script Run on: $($LocalHost) 
"@

Out-File -Append -FilePath $ScriptLog -InputObject $LogMessage

if($Alert) { 
    $Subject = "MicroStrategy Web not available on $IntelligenceServer at $ScriptStartTime"
	$AlertMessage = @" 
Control m job: $($ControlMJobName)
Job Function: $($JobFunction)
"@

    $AlertMessage += $LogMessage

    Send-MailMessage -To $To -Subject $Subject -From $From -SmtpServer $SmtpServer -Body $AlertMessage 
}