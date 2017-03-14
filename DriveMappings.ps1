$DateTime = Get-Date -Format g
$DriveList = @("Mapped Network Drives as of: $DateTime","","Drive    Path","-----    ----")
$MappedDisks = Get-WmiObject -Class Win32_MappedLogicalDisk
$MappedDisks | foreach {
    $DriveName = $_.Name
    $DrivePath = $_.ProviderName
    $DriveList += "$DriveName       $DrivePath"
    }
$DriveList | Out-File H:\mappings.txt
