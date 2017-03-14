# Define column headers for results array.
$results = @("HostName,TestStatus,IPaddress,StatusCode,ResponseTime")
# Get list of servers to test.
$inputs = Get-Content "$PSScriptRoot\input.txt"
# Ping each server to determine if it is online.
$count = 0
Foreach ($server in $inputs)
    {$SrvrAddress = "Address='$server'"
    # Creates custom object to hold results of ping test against current server.
    $MyObj = Get-WmiObject Win32_PingStatus -filter $SrvrAddress | Select address,protocoladdress,statuscode,responsetime
    # Separate ping results into useful fields.
    $hostname = $MyObj.address
    $IPaddress = $MyObj.protocoladdress
    $statuscode = $MyObj.statuscode
    $responsetime = $MyObj.responsetime
    $count = $count + 1
    $count
    # Check "statuscode" and determine if server is online, set "TestStatus" based on result.
    if ($statuscode -ne 0)
    {$TestStatus = "Offline"
    Write-Verbose -Message "Group `"$server`" offline.`n" -Verbose
    $IPaddress = $null
    $statuscode = $null
    $responsetime = $null} Else {
    $TestStatus = "Online"
    $domain = nslookup $IPaddress | Select-String Name
    $hostname = $domain.ToString()
    $hostname = $Hostname.TrimStart("Name:    ")}
    # Write reuslt variables to results array.
    $results += "$hostname,$TestStatus,$IPaddress,$statuscode,$responsetime"
    # Empty custom object for next server.
    $MyObj = $null}
# Write results array to temporary text file.
$results | Out-File "$PSScriptRoot\output.txt"
# Convert temp text file to CSV.
Import-Csv $PSScriptRoot\output.txt | Export-Csv $PSScriptRoot\output.csv -NoTypeInformation
# Delete temp text file.
rm $PSScriptRoot\output.txt
# Cleanup initial arrays.
$results = $null
$inputs = $null
