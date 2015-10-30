#LIST ALL USER CONNECTIONS "BSM_0AP";
[xml]$t = gc bsm*xml                                                      
$cn = $t.CommandManagerResults.ListUserConnection.ChildNodes
$sid = $cn | %{if($_.TimeConnected -gt 600) {$_.SessionId} }                   
$s = "DISCONNECT USER SESSIONID " + [string]::join(", ", $sid)
$s
