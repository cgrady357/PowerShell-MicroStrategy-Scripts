    ##--------------------------------------------------------------------------
    ##  FUNCTION.......:  Create-MicroStrategyGroup
    ##  PURPOSE........:  Creates Command Manager script that will create MicroStrategy user groups and apply security roles to them.
    ##  REQUIREMENTS...:  PowerShell v2
    ##  NOTES..........:  
    ##--------------------------------------------------------------------------
    Function Create-MicroStrategyGroup {
        <#
        .SYNOPSIS
         Creates Command Manager script that will create MicroStrategy user groups and apply security roles to them.
        .DESCRIPTION
         Creates Command Manager script that will create MicroStrategy user groups and apply security roles to them.
        .PARAMETER ProjectName
         The MicroStrategy project name used in script
        .EXAMPLE
         C:\\\\PS>Create-MicroStrategyGroup "Spectrum Dashboard"
         
         This will create Command Manager script that will create MicroStrategy user groups and apply security roles to them.
         
        .NOTES
         NAME......:  Create-MicroStrategyGroup
         AUTHOR....:  Craig Grady
         LAST EDIT.:  04/08/2014
         CREATED...:  04/08/2014
        #>
        
        
        Param(
		[Parameter(
            ValueFromPipeLine = $false,
            Position = 0)
		]
            [string]$project_name
        )  
		
$parent_group = "..Web Groups (NO USER ACCESS AT THIS LEVEL)"
$group = "Web - " + $project_name

$sec_role = "1.Web Reporter - (1.Web Reporter + OLAP + Report Services)", "2.Web Analyst - (2.Web Analyst + OLAP + Report Services)", "3.Web Professional - (3.Web Professional + OLAP + Report Services)", "Architect - (1, 2, 3, 4, 7, 9, 10, 11)"
$priv = "(1.Web Reporter + OLAP + Report Services)", "(2.Web Analyst + OLAP + Report Services)", "(3.Web Professional + OLAP + Report Services)","(Restricted Access)"

$q = '"'
($s = "CREATE USER GROUP " + $q + $group + " (NO USER ACCESS AT THIS LEVEL)" + $q  + " IN GROUP " + $q + $parent_group + $q + ";")
for($c=0; $c -le 3; $c++) {
 $g = $q + $group + " " + $priv[$c] + $q;
 "CREATE USER GROUP " + $g + " IN GROUP " + $q + $group + " (NO USER ACCESS AT THIS LEVEL)" + $q + ";"; 
 "GRANT SECURITY ROLE " + $q + $sec_role[$c] + $q + " GROUP " + $g + " FOR PROJECT " + $q + $project_name + $q + ";";
 }


}