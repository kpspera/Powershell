######################################################################################
###          WINDOWS LOCAL ADMINISTRATOR EVIDENCE COLLECTION SCRIPT                ###  
###       Written by: Kenneth Spera, CISA      Last Update: 03/01/2017             ###
######################################################################################


#################################################################################
###                     Function: Get-LocalAdmins                             ###
###    Collects member information from device local Administrators group.    ###
#################################################################################

function Get-LocalAdmins ($deviceGroup)
    { 
    # Assign valriable values needed to access local groups
    
    $grpList = @()
    $group = [ADSI]("WinNT://$deviceGroup,group")

    # Attempt to obtain group membership and capture errors.

    $Member = @($group.psbase.Invoke("Members")) 2> $null
    if ($error[0] -ne $null)
        {
        $errItem = $error[0].Exception.Message.ToString()
        $errStart = $errItem.IndexOf(":")+1
        $errItem = $errItem.substring($errStart)
        $errItem = $errItem -replace '"',''
        $errLen = $errItem.Length-2
        $errItem = $errItem.substring(0,$errLen)
        $Global:errList += "$Global:server --- $errItem"
        Write-Verbose -message "$Global:server --- $errItem" -Verbose
        $error.clear()
        $grpList = $null
        Return $grpList
        Continue
        }

    # Loop through each local user found and obtain user information.

    foreach ($m in $Member)
        {
        # Get member ID (name)
        
        $MemID = $m.GetType().InvokeMember("Name",'GetProperty',$Null,$m,$Null)
    
        # Identify if account name could not be resolved.

        if ($MemID -like "S-?-?-*")
            {
            $MemDspName = "User account no longer exists. Cannot resolve user name."
            $Global:UsrList += "`"$MemID`",`"`",`"`",`"`",`"$MemDspName`",`"`",`"`""
            $MemDspName = $null
            }

        # If account name resolves, process account.

        Else
            {
            $MemDomain = $m.GetType().InvokeMember("Adspath",'GetProperty',$Null,$m,$Null)
            $MemClass = $m.GetType().InvokeMember("Class",'GetProperty',$Null,$m,$Null)
            $MemClass = $MemClass.ToLower()
            $MemHeir = $deviceGroup

            # If member is a group, prepare for passing to "Get-GroupMembers" function.

            if ($MemClass -eq "group")
                {
                $grpDomain = $MemDomain -match "//.*/"
                $grpDomain = $matches[0]
                $grpDomain = $grpDomain.substring(0,$grpDomain.Length-1) -replace '//',''
                $grpDomain = $grpDomain.ToLower()
                $matches = $null            
                $Global:UsrList += "`"$MemID`",`"$MemClass`",`"$grpDomain`",`"`",`"$MemDspName`",`"`",`"$MemHeir`""
                $grpList += "$grpDomain\$MemID,$deviceGroup"
                $grpDomain = $null
                }
    
        # For members that are not groups, get additional user properties.

           Else
                { 
        
                # Get DisplayName (full name for users)

                $MemDspName = $m.GetType().InvokeMember("FullName",'GetProperty',$Null,$m,$Null)
        
                # Determine if account is disabled.

                $MemEnabled = $m.GetType().InvokeMember("userflags",'GetProperty',$null,$m,$null)
                if($MemEnabled -ne $null)
                    {
                    $bin = [Convert]::ToString($MemEnabled,2)
                    $binpos = $bin.length-2
                    $binval = $bin.Substring($binpos,1)
                    switch ($binval)
                        {
                        0 {$MemEnabled = "True"}
                        1 {$MemEnabled = "False"}
                        }
                    }
        
                # Handle error if account information not found.

                Else
                    {
                    $MemEnabled = 'N/A - Account not found.'
                    }

                # Find and format member domain and create user list entry, then clear variables for next member.

                $mtchResult = $MemDomain -match "//.*"
                $pthTemp = $matches[0]
                $UsrDomain = $pthTemp -replace '//',''
                $mtchResult = $UsrDomain -match ".*?/"
                $dmnTemp = $matches[0]
                $UsrDomain = $dmnTemp -replace '/',''
                $UsrDomain = $UsrDomain.ToLower()
                $MemDept = 'N/A - Local Account'
                $Global:UsrList += "`"$MemID`",`"$MemClass`",`"$UsrDomain`",`"$MemEnabled`",`"$MemDspName`",`"$MemDept`",`"$MemHeir`""
                $mtchResult=$matches=$pthTemp=$dmnTemp=$UsrDomain=$MemID=$MemDspName=$MemDomain=$MemClass=$MemHeir=$MemDept=$MemEnabled = $null
                }
            }
        }
	
    #Ensure that only list of AD network groups found is returned.

    return $grpList
    }

#################################################################################
###                     Function: Get-GroupMembers                            ###
###          Collects group memeber information from Active Directory         ###
#################################################################################

function Get-GroupMembers ($Groups)
    {
    
    # Process groups identified

    Foreach($g in $Groups)
        {
        # Process group list to isolate domain, name, and parents
        $Usrs = @()
        $usrAccounts = @()
        $gHeirarchy = $g -match ",.*"
        $gHeirarchy = $matches[0]
        $gHeirarchy = $gHeirarchy -replace ',',''
        $g = $g -replace ",.*",''
        $gDomain = $g -match ".*?\\"
        $gDomain = $matches[0]
        $gDomain = $gDomain -replace "\\",''
        $gName = $g -creplace "^[^\\]*\\",''
        
        # Determine correct domain to use in Active Directory search

        switch ($gDomain)
            {
            [subdomainname] {$srchDomain = "[subdomain.domain.tld"}
            ...
            }

        # Get members for group and add to group heirarchy

        $Members = Get-ADGroupMember -Identity $gName -Server $srchDomain | Select distinguishedName,objectClass
        $gHeirarchy = "$gHeirarchy -> $g"
        
        # Process group members
        
        Foreach ($mem in $Members)
            {
            # Determine if member is a group
              
            If ($mem.objectClass -eq "group")
                {
                $gClass = $mem.objectClass
                       
                # Isolate group name

                $grpName = $mem.distinguishedName -match "CN=.*?,"
                $grpName = $matches[0]
                $grpName = $grpName -replace "CN=",''
                $grpName = $grpName -replace ',',''

                # Isolate group domain

                $grpDomain = $mem.distinguishedName -match "DC=.*?,"
                $grpDomain = $matches[0].ToUpper()
                $grpDomain = $grpDomain -replace "DC=",''
                $grpDomain = $grpDomain -replace ',',''
                $grpDomain = $grpDomain.ToLower()

                # Compose group list entry

                $GrpList += "$grpDomain\$grpName,$gHeirarchy"

                # Add to member list

                $uDisp = $null
                $uDep = $null
                $Global:UsrList += "`"$grpName`",`"$gClass`",`"$gDomain`",`"`",`"$uDisp`",`"$udep`",`"$gHeirarchy`""
                }
                
            # ...if not, process as a user.

            Else
                {
                
                $Usrs += $mem.DistinguishedName
                }
            }

                # Get user information from Active Directory

            $usrAccounts = $Usrs | Get-ADUser -Server [subdomain.domain.tld]:3268 -Properties SamAccountName, DisplayName, Department, ObjectClass, DistinguishedName, Enabled | Select SamAccountName, DisplayName, Department, ObjectClass, DistinguishedName, Enabled
                
            ForEach ($u in $usrAccounts)
                {
                $uName = $u.SamAccountName
                $uDisp = $u.DisplayName
                $uDep = $u.Department
                $uClass = $u.ObjectClass
                $uEnabled = $u.Enabled
                $uDName = $u.DistinguishedName

                # Process domain from user information

                $uDName = $uDName -match "DC=.*,"
                $uDName = $matches[0]
                $uDName = $uDName -replace "DC=",''
                $uDName = $uDName -replace ",.*",''

                # Compose user list entry and clear variables for next user.

                $Global:UsrList += "`"$uname`",`"$uClass`",`"$uDName`",`"$uEnabled`",`"$uDisp`",`"$uDep`",`"$gHeirarchy`""
                $uname=$uDisp=$uDep=$uClass=$uEnabled=$uDName=$grpName=$grpDomain = $null
                }
            }
    Return $GrpList
    }

#######################################################################################
###                      Function: ConvertTo-XLStyle                                ###
###              Converts CSV to Excel file and formats fields.                     ###
#######################################################################################

function ConvertTo-XLStyle ($file)
    {
    $xlShiftDown = -4121

    # Setup Excel for conversion

    $xlFixedFormat = [Microsoft.Office.Interop.Excel.XlFileFormat]::xlWorkbookDefault
    $xl = new-object -comobject excel.application
    $xl.visible = $false

    # Open and format file

    $wb = $xl.workbooks.open("$file")
    $ActiveSheet = $wb.ActiveSheet
    $range = $ActiveSheet.UsedRange
    $table = $ActiveSheet.ListObjects.add(1,$range,0,1)
    switch ($xl.Version)
        {
        '15.0' {$table.TableStyle = "TableStyleMedium7"}
        '14.0' {$table.TableStyle = "TableStyleMedium4"}
        }
    $range.EntireColumn.AutoFit() | Out-Null
    
    # Find IDs that do not match naming convention for individual user [pattern]. Highlight non-individual-user
    # accounts. Also highlight any broken SIDs.

    $column = 1
    $row = 1
    $MaxRows = $ActiveSheet.UsedRange.Rows.Count
    $Cells = $ActiveSheet.Cells
    $GenIDRows = @()
    $BknSIDRows = @()
    While ($row -le $MaxRows)
        {
        $row++        
        If ($Cells.item($row,$column).Text -match "^[a-aA-Z]{1}\d{6}$" -or $Cells.item($row,$column).Text -match "^[a-aA-Z]{2}\d{5}$"){Continue}
        ElseIf ($Cells.item($row,2).Text -eq 'user') {$GenIDRows += "$row"}
        ElseIf ($Cells.item($row,$column).Text -like "S-?-?-*") {$BknSIDRows +=  "$row"}
        Else {Continue}
        }
    ForEach ($r in $GenIDRows)
        {
        $GenIDRange = $xl.Range("A${r}:G$r")
        $GenIDRange.Interior.ColorIndex = 6
        }
    $r = $null
    ForEach ($r in $BknSIDRows)
        {
        $BknSIDRange = $xl.Range("A${r}:G$r")
        $BknSIDRange.Interior.ColorIndex = 3
        $BknSIDRange.Font.ColorIndex = 2
        }
    
    # Add Legend and format

    $eRow = $Cells.item(1,1).entireRow
    $eRow.Select() | Out-null
    $eRow.insert($xlShiftDown) | Out-Null
    For ($i=1;$i -le 3;$i++)
        {
        $eRow = $Cells.item(1,1).entireRow
        $eRow.Select() | Out-null
        $eRow.insert($xlShiftDown) | Out-Null
        $Cells.item(1,1).BorderAround(1,2,1) | Out-Null
        $Cells.item(1,2).BorderAround(1,2,1) | Out-Null
        }
    $Cells.item(1,1) = "Legend"
    $Cells.item(1,2) = "Color Code"
    $Cells.item(2,1) = "Generic ID"
    $Cells.item(2,2) = "<Text>"
    $Cells.item(3,1) = "Broken SID/Unresolved account or group."
    $Cells.item(3,2) = "<Text>"
    $Cells.item(1,5) = "Source System: $server"
    $Cells.item(2,5) = "Run Date: $Global:rundate"
    $ActiveRange = $Cells.Range("A1","B1")
    $ActiveRange.Select() | Out-Null
    $ActiveRange.Interior.ColorIndex = 10
    $Cells.Range("A1","B1").Font.ColorIndex = 2
    $Cells.Range("B1","B1").EntireColumn.AutoFit() | Out-Null
    $Cells.item(2,2).Interior.ColorIndex = 6
    $Cells.item(3,2).Interior.ColorIndex = 3
    $Cells.item(3,2).Font.ColorIndex = 2    
    
    # Save file and exit Excel

    $xl.ActiveWorkbook.SaveAs("$PSScriptRoot\$server-AdminMembers.xlsx",$xlFixedFormat)
    $xl.Quit()

    # Remove CSV source file

    rm $file | Out-Null
    }

#####################################################################################################################
###                                         Begin Script Processing                                               ###
#####################################################################################################################

Import-Module ActiveDirectory
Add-Type -AssemblyName Microsoft.Office.Interop.Excel

# Ensure error variable is clear
$error.clear()

# Define variables

$Global:errList = @()
$Groups = @()
$Count = 0

# Import list of servers to test

$inputs = Get-Content "$PSScriptRoot\input.txt"

# Find total count of servers to process

$inCount = $inputs.Count

# Begin processing servers

Foreach ($server in $inputs)
    {
    # Prepare UserList variable and increment counter
    
    $Global:UsrList = @("`"UserName`",`"Class`",`"Domain`",`"Enabled`",`"DisplayName`",`"Department`",`"Heirarchy`"")
    $Count = $Count + 1
    "Processing sample $Count/$inCount...$server"

    # Test if anybody's home (server is reachable). If not, add entry to error list, write notice to console, and skip to next server

    $AnybodyHome = Test-Connection -ComputerName $:server -Count 1 -Quiet
    if ($AnybodyHome -eq $false) {
        Write-Verbose -Message "Server `"$server`" could not be reached." -Verbose
        $Global:errList += "$server -- Could not reach server (Ping fail)."
        Continue
        }

    # If Server is reachable process local admin group membership

    Else
        {
        $deviceGroup = "$server/Administrators"
        $Groups = Get-LocalAdmins $deviceGroup

        # Recursively process nested AD groups to get end users

        while ($Groups -ne $null)
            {
            $Groups = Get-GroupMembers $Groups
            }
        
        # Export user list to file and convert to CSV

        $Global:rundate = Get-Date -Format g
        $Global:UsrList >> $PSScriptRoot\$server-AdminMembers.txt
        If(Test-Path $PSScriptRoot\$server-AdminMembers.csv)
            {
            Remove-Item $PSScriptRoot\$server-AdminMembers.csv
            }
        Import-Csv $PSScriptRoot\$server-AdminMembers.txt | Export-Csv $PSScriptRoot\$server-AdminMembers.csv -NoTypeInformation
        rm $PSScriptRoot\$server-AdminMembers.txt
        
        # Clear User list for next system

        $Global:UsrList = @()
        
        # Convert CSV to XLSX and format
        
        $file = "$PSScriptRoot\$server-AdminMembers.csv"
        ConvertTo-XLStyle ($file) | Out-Null
        If(Test-Path $PSScriptRoot\$server-AdminMembers.xlsx) {}
        Else {Write-Verbose -Message "File did not save correctly. Please check script and try again." -Verbose}
    }
}
# Write Errors to file

$Global:errList | Out-File $PSScriptRoot\AdminTest-Errors.txt

# Terminate script

Write-Verbose -Message "Script has completed.`n`n" -Verbose
Pause
Exit
