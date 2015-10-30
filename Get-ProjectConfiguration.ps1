Function Get-ProjectConfiguration {
        <#
        .DESCRIPTION
         Creates and executes a MicroStrategy Command Manager script that will get project configuration
        .PARAMETER ServerName
		 The i-server name you are connecting Command Manager to
        .PARAMETER UserName
		 The user name that will connect to the i-server
		.PARAMETER Password
		 The password used to connect to the i-server
		.PARAMETER Script
		 The name and path of the Command Manager script
		.PARAMETER LogFile
		 The file that Command Manager will log its output to
		.EXAMPLE 
         C:\\\\PS> Get-ProjectConfiguration -ServerName server_name -UserName user -Password my_password -Script "C:\script.scp" -LogFile "C:\log.txt"
         
        .NOTES
         NAME......:  Get-ProjectConfiguration
         AUTHOR....:  Craig Grady
         LAST EDIT.:  11/24/2014
         CREATED...:  11/24/2014
        #>
       
        Param(
		    [string][] $ProjectName,
			[string] $ServerName,
			[string] $UserName,
			[string] $Password,
			[string] $Script,
			[string] $LogFile
        )  
		

$input = Get-Content $InputFile		
$q = '"'

[string] $s = $null
for($c=0; $c -le $ProjectName.count -1; $c++) {
$s += "LIST ALL PROPERTIES FOR PROJECT CONFIGURATION IN PROJECT " + $q + $input[$c][0] + $q + ";" + [Environment]::Newline;
}

Out-File -FilePath $Script -InputObject $s

cmdmgr -n $ServerName -u $UserName -p $Password -f $Script -o $LogFile

}
