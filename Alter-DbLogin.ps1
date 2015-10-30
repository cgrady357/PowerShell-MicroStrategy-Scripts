##-------------------------------------------------------------------------------
    ##  FUNCTION.......:  Alter-DbLogin
    ##  PURPOSE........:  Creates a Command Manager script that will alter a Db Login.
    ##-------------------------------------------------------------------------------
	
    Function Alter-DbLogin {
        <#
        .DESCRIPTION
         Creates a Command Manager script that will alter a Db Login.
        .PARAMETER DbLogin
         The Db Login name
        .OPTIONAL PARAMETER Name
         The new dblogin name
        .OPTIONAL PARAMETER Login
         The new database login 
        .PARAMETER Password
         The new database password
        .EXAMPLE
         C:\\\\PS>Alter-DbLogin -DbLogin "my_db_login_name" -Password "my_new_db_login_password"
         
        .NOTES
         NAME......:  Alter-DbLogin
         AUTHOR....:  Craig Grady
         LAST EDIT.:  10/22/2014
         CREATED...:  10/22/2014
        #>
        
        [CmdletBinding()]
        Param(
		    [Parameter(Mandatory=$true)]
		    [string]$DbLogin,
			[string]$Name,
			[string]$Login,
			[string]$Password
        )  

		
$q = '"'		
$cmd = "ALTER DBLOGIN " + $q + $DbLogin + $q
if($Name) {$cmd += " NAME " + $q + $Name + $q}
if($Login) {$cmd += " LOGIN " + $q + $Login + $q}
if($Password) {$cmd += " PASSWORD " + $q + $Password + $q}

$cmd += ";"
$cmd
}