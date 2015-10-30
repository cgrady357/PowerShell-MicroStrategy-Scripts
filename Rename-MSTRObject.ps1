Function Rename-MSTRObject {
<#
		.NAME
		 Rename-MSTRObject
		.DESCRIPTION
         Renames a list of MicroStrategy objects
        .INPUT
		 Uses an export to text of a MicroStrategy Search for Objects.  Assumes this format "Name"\t"Type"\t"Modification Time"\t"Path"
		.EXAMPLE 
         C:\\\\PS> Rename-MSTRObject -InputFile "C:\my_input.txt" -Project "TN Redet MAXDAT DEV" -OldText "Application Request/Material Request" -ReplacementText "Process Letters" -ScriptName "C:\script.scp" 
		.NOTES
					
         AUTHOR....:  Craig Grady
         LAST EDIT.:  10/30/2015
         CREATED...:  10/29/2015
		 
#>
		 
        Param(
		[Parameter(
            ValueFromPipeLine = $false,
            Position = 0)
		]
			[string] $InputFile,
			[string] $Project,
			[string] $OldText,
			[string] $ReplacementText,
			[string] $ScriptName
        )  
		 

Get-Content $InputFile | Select-Object -Skip 1  | 
%{$Name, $Type, $Path = ($_ -split "\t")[0,1,3]; 
$Lookup_Name = $Name -replace "`""; $Path = $Path -replace ("\\$Project|\\$Lookup_Name"); $Type = $Type.ToUpper() -replace "`"", ""; 
if($Type -eq "Grid"){$Type = "REPORT"};$New_Name = $Name -replace $OldText, $ReplacementText; 
if($Type -eq "Filter"){$op = "ON"} else {$op = "FOR"}; 
"ALTER $Type $NAME IN FOLDER $Path NAME $New_Name $op PROJECT `"$Project`";" }| 
Out-File -FilePath $ScriptName -Encoding ascii
}