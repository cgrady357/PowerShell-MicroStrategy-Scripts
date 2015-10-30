<#
        .NAME
		 Check-MicroStrategyServices
		.DESCRIPTION
         Checks that the MicroStrategy Intelligence Server, MicroStrategy Listener, MicroStrategy Health Agent, and MicroStrategy Enterprise Manager Data Loader services are running.  It will send an email alert if they aren't running.
        .PARAMETER MicroStrategyServices
		 An array of MicroStrategy Services checked to see if they are running
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
         CREATED...:  08/12/2014
#>
        
[CmdletBinding()]
Param(
[string[]]$MicroStrategyServices = ("MicroStrategy Intelligence Server", "MicroStrategy Listener", "MicroStrategy Health Agent", "MicroStrategy Enterprise Manager Data Loader"),
[Parameter(Mandatory=$true)]
[string]$SmtpServer,
[Parameter(Mandatory=$true)]
[string]$From,
[Parameter(Mandatory=$true)]
[string]$To,
[Parameter(Mandatory=$true)]
[string]$ControlMJobName,
[string]$JobFunction = "Checks that the MicroStrategy Intelligence Server, MicroStrategy Listener, MicroStrategy Health Agent, and MicroStrategy Enterprise Manager Data Loader services are running.  It will send an email alert if they aren't running.",
[Parameter(Mandatory=$true)]
[string]$ScriptLogPath,
[string]$ScriptLog = $ControlMJobName + (get-date -f "-yyyy-MM-dd") + ".log"
)

$ScriptStartTime = Get-Date
$Server = $env:ComputerName
$nl = [Environment]::NewLine
$Alert = $ServiceDown = $ServiceHung = $false
$ServicesDownList = New-Object System.Collections.ArrayList
$ServicesNotRespondingList = New-Object System.Collections.ArrayList

$Message += $ScriptStartTime + " MicroStrategy Services Check on server " + $Server

foreach ($Item in $MicroStrategyServices) { 
   $Svc = Get-Service -DisplayName $Item 
   $Message += $Svc.DisplayName + " " + $Svc.Status + $nl; 
   if($Svc.Status -ne "Running"){ $Alert = $true; $ServiceDown = $true; $ServicesDownList.Add($Svc.DisplayName) } 
   else {
     $Service = Get-WmiObject -Class Win32_Service -Filter "Name='$($svc.Name)'"
     $Process = Get-Process -Id $Service.processID
	 if(!$Process.Responding) {  $Alert = $true; $ServiceHung = $true; $ServicesNotRespondingList.Add($Svc.DisplayName) }
   } 
}
$Message.Trim()
if($ServiceDown) { $Message += "Services Down: " + $($ServicesDownList -join ", ") + $nl + "   "}
if($ServiceHung) { $Message += "Services Non-responsive: " + $($ServicesNotRespondingList -join ", ") }

Out-File -Append -FilePath $ScriptLog -InputObject $Message

if($Alert) {
    $Subject = "On server " + $Server + " "
    if($ServiceDown) { $Subject += "MicroStrategy Services Down: " + $($ServicesDownList -join ", ") }
	if($ServiceHung) { $Subject += "MicroStrategy Services Hung: " + $($ServicesNotRespondingList -join ", ") }
    Send-MailMessage -To $To -Subject $Subject -From $From -SmtpServer $SmtpServer -Body $Message 
}