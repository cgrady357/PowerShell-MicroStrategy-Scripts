Function Execute-CommandManager {
<#

		.NAME
		 Execute-CommandManager
		.DESCRIPTION
         Executes a MicroStrategy Command Manager Script.
        .NOTES
		 			
         AUTHOR....:  Craig Grady
         LAST EDIT.:  01/29/2015
         CREATED...:  01/12/2015
		 
		 .Syntax
		 MicroStrategy Command syntax
		Usage: Graphical Interface of Command Manager
		cmdmgrw

		Usage: Interactive Command Line
		cmdmgr -interactive

		Usage: MicroStrategy Metadata
		cmdmgr -n Project_Source_Name [-u Username] [-p password] -f InputFile [-o OutputFile | -break | -or ResultsFile -of FailFile -os SuccessFile] [-xml XMLFile] [-i] [-h] [-showoutput] [-stoponerror] [-e] [-suppresshidden]

		Usage: MicroStrategy Metadata Connection-less
		cmdmgr -connlessmstr -f InputFile [-o OutputFile | -break | -or ResultsFile -of FailFile -os SuccessFile] [-xml XMLFile] [-i] [-h] [-showoutput] [-stoponerror] [-e] [-suppresshidden]

		Usage: Narrowcast Server
		cmdmgr -w ODBC_DSN [-u Login] [-p password] [-d database] [-s SystemPrefix] -f InputFile [-o OutputFile | -break | -or ResultsFile -of FailFile -os SuccessFile] [-xml XMLFile] [-i] [-h] [-showoutput] [-stoponerror] [-e] [-suppresshidden]

		Usage: NarrowCast Server Connection-less
		cmdmgr -connlessNCS [-d database] [-s SystemPrefix] -f InputFile [-o OutputFile | -break | -or ResultsFile -of FailFile -os SuccessFile] [-xml XMLFile] [-i] [-h] [-showoutput] [-stoponerror] [-e] [-suppresshidden]

		Parameters enclosed in brackets ("[" & "]") are optional.
		-n              Project Source Name
		-w              (Narrocast Server) Data Source Name:
		-d              (Narrocast Server) Database Name
		-s              (Narrocast Server) System Prefix
		-u              Login (not required for sources configured for Windows
						authentication)
		-p              Password
		-connlessmstr   MicroStrategy Metadata MicroStrategy Metadata Connection-less
		-connlessncs    NarrowCast Server MicroStrategy Metadata Connection-less
		-f              Script input file
		-i              Include Instructions in the log file(s)
		-h              Include file log header
		-showoutput     Displays output on the console
		-suppresshidden Suppress hidden object(s) in the results
		-e              Include error codes in the log file(s)
		-stoponerror    Stop script execution on error
		-break          Separate the ouput into three files: CmdMgrSuccess.log,
						CmdMgrFail.log, CmdMgrResults.log
		-os             Success log file
		-of             Error log file
		-or             Results log file
		-o              Output log file
		-xml            Output XML file
		Output options ((-break | -o | -or -of -os)) are mutually exclusive.

		Note: Project Source Names and file names cannot contain spaces unless the names are enclosed in double quotes.

#>

[CmdletBinding()]
param(
[Parameter(Mandatory=$true)]
[string]$ProjectSourceName,
[string]$UserName,
[string]$Password,
[Parameter(Mandatory=$true)]
[string]$InputFile,
[Parameter(ParameterSetName='Output')]
[string]$OutputFile,
[Parameter(ParameterSetName='Break')]
[Switch]$Break,
[Parameter(ParameterSetName='Results')]
[string]$ResultsFile,
[Parameter(ParameterSetName='Results')]
[string]$FailFile,
[Parameter(ParameterSetName='Results')]
[string]$SuccessFile,
[string]$XMLFile,
[Switch]$Instructions,
[Switch]$LogHeader,
[Switch]$ShowOutputOnConsole,
[Switch]$StopOnError,
[Switch]$IncludeErrorCodes,
[Switch]$SuppressHidden
)

$Empty = [String]::Empty
$Cmd = ("E:\Program Files (x86)\MicroStrategy\Command Manager\CmdMgr.exe", "E:\Program Files\MicroStrategy\Command Manager\CmdMgr.exe" | %{ if(Test-Path $_) { $_ }} | Select-Object -first 1)
$ArgList = New-Object System.Collections.ArrayList
switch ($true) { 
  ($ProjectSourceName -ne $Empty) 	{ $ArgList.Add("-n"); $ArgList.Add($ProjectSourceName)}   
  ($UserName -ne $Empty)			{ $ArgList.Add("-u"); $ArgList.Add($UserName)}   
  ($Password  -ne $Empty)			{ $ArgList.Add("-p"); $ArgList.Add($Password)}   
  ($InputFile  -ne $Empty)			{ $ArgList.Add("-f"); $ArgList.Add($InputFile)}   
  ($OutputFile  -ne $Empty)			{ $ArgList.Add("-o"); $ArgList.Add($OutputFile)}   
  $Break							{ $ArgList.Add("-break")}   
  ($ResultsFile  -ne $Empty)		{ $ArgList.Add("-or"); $ArgList.Add($ResultsFile)}   
  ($FailFile   -ne $Empty)			{ $ArgList.Add("-of"); $ArgList.Add($FailFile)}   
  ($SuccessFile   -ne $Empty)		{ $ArgList.Add("-os"); $ArgList.Add($SuccessFile)}   
  ($XMLFile   -ne $Empty)			{ $ArgList.Add("-xml"); $ArgList.Add($XMLFile)}   
  $Instructions						{ $ArgList.Add("-i")}   
  $LogHeader						{ $ArgList.Add("-h")}   
  $ShowOutputOnConsole				{ $ArgList.Add("-showoutput")}   
  $StopOnError						{ $ArgList.Add("-stoponerror")}   
  $IncludeErrorCodes				{ $ArgList.Add("-e")}   
  $SuppressHidden					{ $ArgList.Add("-suppresshidden")}   
}

& $Cmd $ArgList
}