  ##--------------------------------------------------------------------------
    ##  FUNCTION.......:  Create-FolderACE
    ##  PURPOSE........:  Creates a Command Manager script that will add ACEs to \Public Folders and \Schema Objects 
    ##  REQUIREMENTS...:  PowerShell v2
    ##  NOTES..........:  
    ##--------------------------------------------------------------------------
    Function Create-FolderACE {
        <#
        .SYNOPSIS
         Creates a Command Manager script that will add ACEs to \Public Folders and \Schema Objects 
        .DESCRIPTION
         Creates a Command Manager script that will add ACEs to \Public Folders and \Schema Objects 
        .PARAMETER ProjectName
         The MicroStrategy project name used in script
        .EXAMPLE
         C:\\\\PS>Create-FolderACE "Spectrum Dashboard"
         
         This will create a Command Manager script that will add ACEs to \Public Folders and \Schema Objects 
         
        .NOTES
         NAME......:  Create-FolderACE
         AUTHOR....:  Craig Grady
         LAST EDIT.:  04/10/2014
         CREATED...:  04/10/2014
        #>
        
        
        Param(
		[Parameter(
            ValueFromPipeLine = $false,
            Position = 0)
		]
            [string]$ProjectName,
			[string]$UserGroup
        )  
		

$group = "Web - " + $UserGroup
$priv = "(1.Web Reporter + OLAP + Report Services)", "(2.Web Analyst + OLAP + Report Services)", "(3.Web Professional + OLAP + Report Services)","(Restricted Access)"
$folder = "Public Objects", "Schema Objects"
$q = '"'

$f=0
for($c=0; $c -le 3; $c++) {
 $g = $q + $group + " " + $priv[$c] + $q;
 "ADD ACE FOR FOLDER " + $q + $folder[$f] + $q + " IN FOLDER " + $q + "\" + $q + " GROUP " + $g + " ACCESSRIGHTS DEFAULT CHILDRENACCESSRIGHTS  FULLCONTROL FOR PROJECT " + $q + $ProjectName + $q + ";"; 
 }
 "ALTER ACL FOR FOLDER " + $q + $folder[$f] + $q + " IN FOLDER " + $q + "\" + $q + " PROPAGATE RECURSIVELY FOR PROJECT " + $q + $ProjectName + $q + ";"; 

# only give schema objects full control to architects 
$f++
 "ADD ACE FOR FOLDER " + $q + $folder[$f] + $q + " IN FOLDER " + $q + "\" + $q + " GROUP " + $q + $group + " " + $priv[-1] + $q + " ACCESSRIGHTS DEFAULT CHILDRENACCESSRIGHTS FULLCONTROL FOR PROJECT " + $q + $ProjectName + $q + ";"; 
 "ALTER ACL FOR FOLDER " + $q + $folder[$f] + $q + " IN FOLDER " + $q + "\" + $q + " PROPAGATE RECURSIVELY FOR PROJECT " + $q + $ProjectName + $q + ";"; 


}
