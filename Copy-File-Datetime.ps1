<#
        .NAME
		 Copy-File-DateTime
		.DESCRIPTION
         Copies a file to add the current datetime to the file name
        .PARAMETER Path
	     The file name and full path of the file you want to add the current datetime to the name
		
        .NOTES
         AUTHOR....:  Craig Grady
         LAST EDIT.:  01/29/2015
         CREATED...:  01/29/2015
#>

[CmdletBinding()]
param(
[Parameter(Mandatory=$true)]
[string]$Path
)

if (!(Test-Path $Path)) { "Invalid Path"; exit }

$File = Get-ChildItem $Path
$NewPath = (($File.PSParentPath -split("::"))[1]) + "\" + $File.BaseName + (Get-Date -Format '-yyyy-MM-dd_H-m-s') + $File.Extension 
Copy-Item -Path $Path -Destination $NewPath


