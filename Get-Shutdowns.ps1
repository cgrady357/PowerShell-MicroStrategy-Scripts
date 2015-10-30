cd "E:\Program Files\Common Files\MicroStrategy\Log"
$log = "H:\" + $env:ComputerName + "_shutdowns.txt"
get-content .\DSSErrors.log* | % { if ($_ -match "shutdown") { $_ } } > $log
