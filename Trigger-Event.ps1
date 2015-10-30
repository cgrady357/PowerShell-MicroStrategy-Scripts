<#
        .NAME
		 Trigger-Event
		.DESCRIPTION
         Triggers a MicroStrategy Intelligence Server event using MicroStrategy Command Manager.
		.PARAMETER ControlMJobName 
		 The name of the control m job used to kick off script 
        .PARAMETER EventName
		 The name of the MicroStrategy Intelligence Server event that will be triggered
       
		.Example  
		 powershell "C:\ProgramData\Trigger-Event.ps1" -ControlMJobName "m0apdtr33" -EventName "Daily_RISE_Load"
        
        .NOTES
         AUTHOR....:  Craig Grady
         LAST EDIT.:  01/06/2015
         CREATED...:  01/05/2015
#>

[CmdletBinding()]
param(
[Parameter(Mandatory=$true)]
[string]$ControlMJobName,
[Parameter(Mandatory=$true)]
[string]$EventName
)

[string]$ProjectSourceName = $env:ComputerName 
[string]$UserID,
[string]$Password,
[string]$SmtpServer,
[string]$From, 
[string]$To, 
[string]$JobFunction = "Triggers a MicroStrategy Intelligence Server event using MicroStrategy Command Manager"
[string]$ScriptLogPath = "E:\Prod_E\0AP\Logs\MicroStrategy"
[string]$ScriptLog = $ControlMJobName + $(get-date -f "-yyyy-MM-dd") + ".log" 
$ScriptLog = (Join-Path -Path $ScriptLogPath -ChildPath $ScriptLog)
[string]$CmdMgrScript = Join-Path -Path $ScriptLogPath -ChildPath ($ControlMJobName.Substring(4.5) + ".scp")
[string]$CmdMgrLogFile = Join-Path -Path $ScriptLogPath -ChildPath ($ControlMJobName.Substring(4.5) + ".log")

$ScriptStartTime = Get-Date
$nl = [Environment]::NewLine


$Command = "TRIGGER EVENT `"$EventName`";";

Out-File -FilePath $CmdMgrScript -InputObject $Command

cmdmgr -n $ProjectSourceName -u $UserID -p $Password -f $CmdMgrScript -o $CmdMgrLogFile

$Status = Get-Content $CmdMgrLogFile | %{ if ($_ -match "Event") {$_} }

if($Status -match "success") {$State = "Success"} else {$State = "Failure"}
$Subject = $State + ": " + $EventName + " Event Triggered in MicroStrategy Intelligence Server on " + $env:ComputerName 

Send-MailMessage -To $To -Subject $Subject -From $From -SmtpServer $SmtpServer -Body $status 

Remove-Item $CmdMgrScript 
$NewFileName = (Join-Path -Path $ScriptLogPath -ChildPath $ControlMJobName) + $ScriptStartTime.ToString("-yyyy-MM-dd-hh-mm") + ".log"
Rename-Item -Path $CmdMgrLogFile -NewName $NewFileName