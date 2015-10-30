$path = "HKLM:\SOFTWARE\Wow6432Node\ODBC\ODBC.INI"
Get-ChildItem -Path $path -Exclude "Default", "ODBC Data Sources" | 
foreach{"Name: $($_.ToString().Split("\")[-1])"; foreach($p in $_.Property) {"$($p): $($_.GetValue($p))"}"`n" }
