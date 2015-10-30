Function Alter-Metric {
        <#
        .DESCRIPTION
         Creates and executes a MicroStrategy Command Manager script that will alter metric objects 
        .PARAMETER InputFile
         The file containing the data for the Command Manager script - Metric Name, Metric Path, Description, Long Description, Project Name
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
         C:\\\\PS> Alter-Metric -InputFile "C:\my_input.txt" -ServerName server_name -UserName user -Password my_password -Script "C:\script.scp" -LogFile "C:\log.txt"
         
        .NOTES
         NAME......:  Alter-Metric
         AUTHOR....:  Craig Grady
         LAST EDIT.:  06/16/2014
         CREATED...:  06/16/2014
        #>
       
        Param(
		[Parameter(
            ValueFromPipeLine = $false,
            Position = 0)
		]
			[string] $InputFile,
			[string] $ServerName,
			[string] $UserName,
			[string] $Password,
			[string] $Script,
			[string] $LogFile
        )  
		

$input = Get-Content $InputFile		
$q = '"'

[string] $s = $null
for($c=0; $c -le $input.count -1; $c++) {
$s += "ALTER METRIC " + $q + $input[$c][0] + $q + " IN FOLDER " + $input[$c][1] + $q + " DESCRIPTION " + $q + $input[$c][2] + $q + " LONGDESCRIPTION " + $q + $input[$c][3] + $q + " FOR PROJECT " + $q + $input[$c][4] + $q + ";" + [Environment]::Newline;
}

Out-File -FilePath $Script -InputObject $s

cmdmgr -n $ServerName -u $UserName -p $Password -f $Script -o $LogFile

}
