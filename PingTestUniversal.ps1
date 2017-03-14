$Count = 0
$pingresults = @("ServerName,OnlineStatus,IPAddress")
$inputs = Get-Content "$PSScriptRoot\input.txt"
Foreach ($server in $inputs)
    { $Count = $Count + 1
    "Processing sample $Count..."
    $pings = ping $server /n 1
    if ($pings[2] -like 'Request timed out.')
        {$onlinestatus = "Potentially Offline/Response timed out"
        Write-Verbose -Message "Server `"$server`" potentially offline or request timed out.`n" -Verbose
        $pings = $pings[1] | % { if($_.IndexOf("[") -gt 0) {$_.split("]")[0]}}
        $pings = $pings.split('[')[-1]
        $FQDN = nslookup $server | Select -First 1 -Skip 3
        $FQDN = $FQDN -replace "Name:    ",""
        $pingresults += "$FQDN,$onlinestatus,$pings"
        }
    elseif($pings -Like 'Pinging*')
        {
        $pings = $pings[2] | % { if($_.IndexOf(" from ") -gt 0) {$_.Split(":")[0]}}
        $pings = $pings -replace "Reply from ",""
        $onlinestatus = "Online"
        $FQDN = nslookup $server | Select -First 1 -Skip 3
        $FQDN = $FQDN -replace "Name:    ",""
        $pingresults += "$FQDN,$onlinestatus,$pings"
        }
    else
        {
        Write-Verbose -Message "Server `"$server`" could not be found.`n" -Verbose
        $onlinestatus = "Offline/Not Found"
        $pingresults += "$server,$onlinestatus,$pings"
        }
    }
$pingresults | Out-File "$PSScriptRoot\output-PingTestAll.txt"
# Convert temp text file to CSV.
If(Test-Path $PSScriptRoot\output-PingTestAll.csv)
    {Remove-Item $PSScriptRoot\output-PingTestAll.csv}
Import-Csv $PSScriptRoot\output-PingTestAll.txt | Export-Csv $PSScriptRoot\output-PingTestAll.csv -NoTypeInformation
# Delete temp text file.
rm $PSScriptRoot\output-PingTestAll.txt
"`nTest complete. See results in:`n"
"$PSScriptRoot\output-PingTestAll.csv`n"
pause
Clear-Host
