#by Craig Grady
Param(
		[string]$Log = "C:\Program Files\Common Files\MicroStrategy\Log\DSSErrors.log",
		[string]$SeenFile = "C:\Program Files\Common Files\MicroStrategy\Log\seen.txt",
		[string]$SubsToProjs = "C:\Program Files\Common Files\MicroStrategy\Log\subscriptions_to_project_mapping.txt", # file format:  subscription=project 
		[string]$FailedSubsFile = "C:\Program Files\Common Files\MicroStrategy\Log\failed_subs.txt",
		[string]$Script = "C:\Program Files\MicroStrategy\Command Manager\subs.scp",
		[string]$Project_Source = "Server1",
		[string]$Username = "administrator",
		[string]$Password = "my_password",
		[string]$Cmdmgr_log = "C:\Program Files\MicroStrategy\Command Manager\subs.out"
        ) 
Get-Content $Log |  %{ if($_ -match "FAILED SUBSCRIPTION"){ $_ } } | Out-File $FailedSubsFile
$seen = Get-Content $SeenFile
$proj_hash = Get-Content $SubsToProjs | ConvertFrom-StringData
$failed_subs = Get-Content $FailedSubsFile | %{if($_ -notin $seen) { $_ }}
$q= '"'
$failed_subs | %{ $_.split("'")[1] } | %{"TRIGGER SUBSCRIPTION " + $q + $_ + $q + " FOR PROJECT " + $q + $proj_hash.$_ + $q + ";" } | Out-File -Append $Script

cmdmgr -n $Project_Source -u $Username -p $Password -f $Script -o $Cmdmgr_log 

$FailedSubsFile | Out-File $SeenFile