cd "E:\Program Files\Common Files\MicroStrategy\Log"
$log = "H:\" + $env:ComputerName + "_startups.txt"
get-content .\DSSErrors.log* | % { if ($_ -match "StartUpManager" ) { $_ } } > $log
