    ##--------------------------------------------------------------------------
    ##  FUNCTION.......:  Create-MicroStrategyUser
    ##  PURPOSE........:  Creates Command Manager script that will create MicroStrategy users and assign to user groups
    ##  REQUIREMENTS...:  PowerShell v2
    ##  NOTES..........:  
    ##--------------------------------------------------------------------------
    Function Create-MicroStrategyUser {
        <#
        .SYNOPSIS
         Creates Command Manager script that will create MicroStrategy users and assign to user groups
        .DESCRIPTION
         Creates Command Manager script that will create MicroStrategy users and assign to user groups
        .PARAMETER ProjectName
         The MicroStrategy project name used in script
        .EXAMPLE
         C:\\\\PS>Create-MicroStrategyGroup "Spectrum Dashboard"
         
         Creates Command Manager script that will create MicroStrategy users and assign to user groups
         
        .NOTES
         NAME......:  Create-MicroStrategyUser
         AUTHOR....:  Craig Grady
         LAST EDIT.:  04/17/2014
         CREATED...:  04/17/2014
        #>
           
        Param(
		[Parameter(ValueFromPipeLine = $false, Position = 0)] [string]$ProjectName = "Vendor_BTE_EVDO",
		[Parameter(ValueFromPipeLine = $false, Position = 0)] $listbox_input = "C:\Users\cgrady04\Documents\WindowsPowershell\listbox.txt"
        )  
		
Function Get-FileName
{ 
 Param(	[Parameter( ValueFromPipeLine = $false, Position = 0)] [string]$initialDirectory = "C:\Users\cgrady04\Documents\WindowsPowershell"
        )  

 [System.Reflection.Assembly]::LoadWithPartialName("System.windows.forms") |
 Out-Null

 $OpenFileDialog = New-Object System.Windows.Forms.OpenFileDialog
 $OpenFileDialog.initialDirectory = $initialDirectory
 $OpenFileDialog.filter = "All files (*.*)| *.*"
 $OpenFileDialog.ShowHelp = $true
 $OpenFileDialog.ShowDialog() | Out-Null
 $OpenFileDialog.filename
} #end function Get-FileName

$input_users = Get-Content( Get-FileName -initialDirectory "C:\Users\cgrady04\Documents\WindowsPowershell")

$psh = "C:\Users\cgrady04\Documents\WindowsPowerShell"
if(($env:PATH | %{$_.split(";") } | ?{$_ -eq $p }) -eq $null) { $env:PATH += "C:\Users\cgrady04\Documents\WindowsPowerShell;"}
.{.\Get-Listbox.ps1}

$groups = ("Web - " + $ProjectName + " (1.Web Reporter + OLAP + Report Services)"), ("Web - " + $ProjectName + " (2.Web Analyst + OLAP + Report Services)"), ("Web - " + $ProjectName + " (3.Web Professional + OLAP + Report Services)"),  ("Web - " + $ProjectName + " (Restricted Access)")
$selection =  Get-Listbox -listbox_items $groups

$users = @{}
$input_users | %{$name, $login = $_.split("`t"); $first_name, $last_name = $name.split(" "); $users.$login = "$last_name, $first_name"}
$q = '"'
$password = "msi"
foreach($user_login in $users.Keys) {
($s = "CREATE USER " + $q + $user_login + $q + " FULLNAME " + $q + $users.$user_login + $q + "  PASSWORD " + $q + $password + $q + " IN GROUP " + $q + $selection + $q + ";")
}

}