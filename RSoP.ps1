######################################################################################
##                        RESULTANT SET of POLICY XML PARSER                        ##
##                                by Ken Spera, CISA                                ##
##                                                                                  ##
## Last Update: 6/23/2017                                                           ##
##                                                                                  ##
## Arguments:                                                                       ##
## Accempts computername as argument. Ensure that source XML file is named as       ##
## [COMPUTERNAME]-RSOP.xml and is in the same folder as this script.                ##
##                                                                                  ##
## Description:                                                                     ##
## Parses XML-formatted Resutant Set of Policy reports (GPResult) into more useful  ##	
## Excel (.xlsx) format. For each policy returned in the RSoP Report, the policy    ##
## type (Computer vs. User), winning GPO, policy name, state, description,          ##
## supported versions, and all configurations falling under the policy are returned ##
## and formatted to an Excel table.                                                 ##
######################################################################################


$compName = $args[0]
$error.Clear()
Add-Type -AssemblyName Microsoft.Office.Interop.Excel
$xlFixedFormat = [Microsoft.Office.Interop.Excel.XlFileFormat]::xlWorkbookDefault
$xl = new-object -comobject excel.application
$wb = $xl.Workbooks.Add()
$as = $wb.WorkSheets.Item(1)
$xl.Visible = $false
$xml = [xml](Get-Content "$PSScriptRoot\$compName-RSOP.xml")
$ConfigMax = 0

########################################
##        FUNCTION: Format-XL         ##
## Performs basic formatting to EXCEL ##
## Worksheets.                        ##
########################################

Function Format-XL ($xl, $wb, $as) {
$range = $as.UsedRange
$range.WrapText = $true
$cells.EntireColumn.AutoFit() | Out-Null
$atable = $as.ListObjects.add(1,$range,0,1)
switch ($xl.Version)
    {
    '15.0' {$atable.TableStyle = "TableStyleMedium21"}
    '14.0' {$atable.TableStyle = "TableStyleMedium18"}
    }
}

　
#############################################################
## Compile list of all applied Group Policy Objects (GPOs) ##
#############################################################

$GPOs = $xml.DocumentElement.ComputerResults.GPO
$GPList = @()
ForEach ($gp in $GPOs)
    {
    $obj = New-Object PSObject
    $obj | Add-Member -MemberType NoteProperty -Name GPName -Value $gp.Name
    $obj | Add-Member -MemberType NoteProperty -Name GPEnabled -Value $gp.Enabled
    $obj | Add-Member -MemberType NoteProperty -Name GPIsValid -Value $gp.IsValid
    $obj | Add-Member -MemberType NoteProperty -Name GPID -Value $null
    $obj | Add-Member -MemberType NoteProperty -Name GPDomain -Value $null
    $ObjID = $gp | Select Path | Select -ExpandProperty * | Select Identifier | Select -ExpandProperty * | Select `#text
    $obj.GPID = $objID.'#text'
    $GPDomain = $gp | Select Path | Select -ExpandProperty * | Select Domain | Select -ExpandProperty * | Select `#text
    $obj.GPDomain = $GPDomain.'#text'
    $GPList += $obj
    Remove-Variable -Name obj
    }

#############################
## PARSE COMPUTER SETTINGS ##
#############################
$rsop = @()
$CompPolicies = $xml.DocumentElement.ComputerResults.ExtensionData | Select -ExpandProperty extension | Select Policy | Select -ExpandProperty * | Select Category, Name, State, Explain, Supported, EditText, DropDownList, GPO -ErrorVariable CompError -ErrorAction SilentlyContinue
$ResultType = "ComputerResults"

If ($CompError -ne $null)
    {
    $CompError = $Error[0].Exception.Message
    Write-Host -ForegroundColor Red "$CompError"
    $error.Clear()
    }

ForEach ($p in $CompPolicies)
    {
    $num = 0
    $obj = New-Object PSObject
    $obj | Add-Member -MemberType NoteProperty -Name ResultType -Value $ResultType
    $obj | Add-Member -MemberType NoteProperty -Name WinningGPO -Value $null
    $obj | Add-Member -MemberType NoteProperty -Name GPODomain -Value $null
    $obj | Add-Member -MemberType NoteProperty -Name Category -Value $p.Category
    $obj | Add-Member -MemberType NoteProperty -Name Name -Value $p.Name
    $obj | Add-Member -MemberType NoteProperty -Name State -Value $p.State
    $obj | Add-Member -MemberType NoteProperty -Name Explain -Value $p.Explain
    $obj | Add-Member -MemberType NoteProperty -Name Supported -Value $p.Supported

    $PolGPID = $p | Select GPO | Select -ExpandProperty * | Select Identifier | Select -ExpandProperty * | Select `#text
    $GPMatch = $GPList | Where-Object -Property GPID -eq -Value $PolGPID.'#text'
    $obj.WinningGPO = $GPMatch.GPName
    $obj.GPODomain = $GPMatch.GPDomain

　
    If ($p.EditText -ne $null)
        {
        $EditText = $p | Select EditText | Select -ExpandProperty *
        $ETcount = $EditText.Name.Count
        $etTempArray = @()
        ForEach ($et in $EditText)
            { # Creates temporay object to hold parsed results and populates standard information: Name, State, and Value for each configuration
            $etTempObj = New-Object PSObject
            $etTempObj | Add-Member -MemberType NoteProperty -Name etTempName -Value ($et.Name | Out-String).Trim()
            $etTempObj | Add-Member -MemberType NoteProperty -Name etTempState -Value ($et.State | Out-String).Trim()
            $etTempObj | Add-Member -MemberType NoteProperty -Name etTempValue -Value ($et.Value | Out-String).Trim()
            $etTempArray += $ddTempObj
            Remove-Variable -Name etTempObj
            }
        ForEach ($eto in $etTempArray)
            { # Adds all temporary object to the master object list
            $num++
            $bool = [bool]($obj.psobject.Properties | where { $_.Name -eq "CongfigName$num"})
            If ($bool -eq $true)
                {
                $obj."ConfigName$num" = $eto.etTempName
                $obj."ConfigState$num" = $eto.etTempState
                $obj."ConfigValue$num" = $eto.etTempValue                
                }
            Else
                {
                $obj | Add-Member -MemberType NoteProperty -Name ConfigName$num -Value $eto.etTempName
                $obj | Add-Member -MemberType NoteProperty -Name ConfigState$num -Value $eto.etTempState
                $obj | Add-Member -MemberType NoteProperty -Name ConfigValue$num -Value $eto.etTempValue
                }
            If ($num -gt $ConfigMax)
                {
                $rsop | ForEach-Object {$_ | Add-Member -MemberType NoteProperty -Name ConfigName$num -Value $null; $_ | Add-Member -MemberType NoteProperty -Name ConfigState$num -Value $null; $_ | Add-Member -MemberType NoteProperty -Name ConfigValue$num -Value $null}
                $ConfigMax = $num
                }
            $bool = $null
            }
        Remove-Variable -Name etTempArray
        }

    If ($p.DropDownList -ne $null)
        {
        $DropDown = $p | Select DropDownList | Select -ExpandProperty *
        $DDcount = $DropDown.Name.Count
        $ddTempArray = @()
        ForEach ($dd in $DropDown)
            {# Creates temporay object to hold parsed results and populates standard information: Name, State, and Value for each configuration
            $ddTempObj = New-Object PSObject
            $ddTempObj | Add-Member -MemberType NoteProperty -Name ddTempName -Value ($dd.Name | Out-String).Trim()
            $ddTempObj | Add-Member -MemberType NoteProperty -Name ddTempState -Value ($dd.State | Out-String).Trim()
            $ddTempObj | Add-Member -MemberType NoteProperty -Name ddTempValue -Value $null
            $DropDownValue = $dd | Select Value | Select -ExpandProperty *
            $ddTempObj.ddTempValue = ($DropDownValue.Name | Out-String).Trim()
            $ddTempArray += $ddTempObj
            Remove-Variable -Name ddTempObj
            }
        ForEach ($to in $ddTempArray)
            { # Adds all temporary object to the master object list
            $num++
            $bool = [bool]($obj.psobject.Properties | where { $_.Name -eq "CongfigName$num"})
            If ($bool -eq $true)
                {
                $obj."ConfigName$num" = $to.ddTempName
                $obj."ConfigState$num" = $to.ddTempState
                $obj."ConfigValue$num" = $to.ddTempValue
                }
            Else
                {
                $obj | Add-Member -MemberType NoteProperty -Name ConfigName$num -Value $to.ddTempName
                $obj | Add-Member -MemberType NoteProperty -Name ConfigState$num -Value $to.ddTempState
                $obj | Add-Member -MemberType NoteProperty -Name ConfigValue$num -Value $to.ddTempValue
                }
            If ($num -gt $ConfigMax)
                {
                $rsop | ForEach-Object {$_ | Add-Member -MemberType NoteProperty -Name ConfigName$num -Value $null; $_ | Add-Member -MemberType NoteProperty -Name ConfigState$num -Value $null; $_ | Add-Member -MemberType NoteProperty -Name ConfigValue$num -Value $null}
                $ConfigMax = $num
                }
            }
        Remove-Variable -Name ddTempArray
        }
    Else {} 

    $rsop += $obj # adds objec to master list, then removes all temporary variables to prepare for next cycle.
    Remove-Variable -name obj
    Remove-Variable -name p
    Remove-Variable -Name GPMatch
    Remove-Variable -Name PolGPID
    }

　
#########################
## PARSE USER SETTINGS ##
#########################
$UsrPolicies = $xml.DocumentElement.UserResults.ExtensionData | Select -ExpandProperty extension | Select Policy | Select -ExpandProperty * | Select Category, Name, State, Explain, Supported, EditText, DropDownList, GPO -ErrorVariable UsrError -ErrorAction SilentlyContinue
$ResultType = "UserResults"

If ($UsrError -ne $null)
    {
    $UsrError = $Error[0].Exception.Message
    Write-Host -ForegroundColor Red "$UsrError"
    $error.Clear()
    }

ForEach ($p in $CompPolicies)
    {
    $num = 0
    $obj = New-Object PSObject
    $obj | Add-Member -MemberType NoteProperty -Name ResultType -Value $ResultType
    $obj | Add-Member -MemberType NoteProperty -Name WinningGPO -Value $null
    $obj | Add-Member -MemberType NoteProperty -Name GPODomain -Value $null
    $obj | Add-Member -MemberType NoteProperty -Name Category -Value $p.Category
    $obj | Add-Member -MemberType NoteProperty -Name Name -Value $p.Name
    $obj | Add-Member -MemberType NoteProperty -Name State -Value $p.State
    $obj | Add-Member -MemberType NoteProperty -Name Explain -Value $p.Explain
    $obj | Add-Member -MemberType NoteProperty -Name Supported -Value $p.Supported

    $PolGPID = $p | Select GPO | Select -ExpandProperty * | Select Identifier | Select -ExpandProperty * | Select `#text
    $GPMatch = $GPList | Where-Object -Property GPID -eq -Value $PolGPID.'#text'
    $obj.WinningGPO = $GPMatch.GPName
    $obj.GPODomain = $GPMatch.GPDomain

    If ($p.EditText -ne $null)
        {
        $EditText = $p | Select EditText | Select -ExpandProperty *
        $ETcount = $EditText.Name.Count
        $etTempArray = @()
        ForEach ($et in $EditText)
            {
            $etTempObj = New-Object PSObject
            $etTempObj | Add-Member -MemberType NoteProperty -Name etTempName -Value ($et.Name | Out-String).Trim()
            $etTempObj | Add-Member -MemberType NoteProperty -Name etTempState -Value ($et.State | Out-String).Trim()
            $etTempObj | Add-Member -MemberType NoteProperty -Name etTempValue -Value ($et.Value | Out-String).Trim()
            $etTempArray += $ddTempObj
            Remove-Variable -Name etTempObj
            }
        ForEach ($eto in $etTempArray)
            {
            $num++
            $bool = [bool]($obj.psobject.Properties | where { $_.Name -eq "CongfigName$num"})
            If ($bool -eq $true)
                {
                $obj."ConfigName$num" = $eto.etTempName
                $obj."ConfigState$num" = $eto.etTempState
                $obj."ConfigValue$num" = $eto.etTempValue                
                }
            Else
                {
                $obj | Add-Member -MemberType NoteProperty -Name ConfigName$num -Value $eto.etTempName
                $obj | Add-Member -MemberType NoteProperty -Name ConfigState$num -Value $eto.etTempState
                $obj | Add-Member -MemberType NoteProperty -Name ConfigValue$num -Value $eto.etTempValue
                }
            If ($num -gt $ConfigMax)
                {
                $rsop | ForEach-Object {$_ | Add-Member -MemberType NoteProperty -Name ConfigName$num -Value $null; $_ | Add-Member -MemberType NoteProperty -Name ConfigState$num -Value $null; $_ | Add-Member -MemberType NoteProperty -Name ConfigValue$num -Value $null}
                $ConfigMax = $num
                }
            $bool = $null
            }
        Remove-Variable -Name etTempArray
        }

    If ($p.DropDownList -ne $null)
        {
        $DropDown = $p | Select DropDownList | Select -ExpandProperty *
        $DDcount = $DropDown.Name.Count
        $ddTempArray = @()
        ForEach ($dd in $DropDown)
            {
            $ddTempObj = New-Object PSObject
            $ddTempObj | Add-Member -MemberType NoteProperty -Name ddTempName -Value ($dd.Name | Out-String).Trim()
            $ddTempObj | Add-Member -MemberType NoteProperty -Name ddTempState -Value ($dd.State | Out-String).Trim()
            $ddTempObj | Add-Member -MemberType NoteProperty -Name ddTempValue -Value $null
            $DropDownValue = $dd | Select Value | Select -ExpandProperty *
            $ddTempObj.ddTempValue = ($DropDownValue.Name | Out-String).Trim()
            $ddTempArray += $ddTempObj
            Remove-Variable -Name ddTempObj
            }
        ForEach ($to in $ddTempArray)
            {
            $num++
            $bool = [bool]($obj.psobject.Properties | where { $_.Name -eq "CongfigName$num"})
            If ($bool -eq $true)
                {
                $obj."ConfigName$num" = $to.ddTempName
                $obj."ConfigState$num" = $to.ddTempState
                $obj."ConfigValue$num" = $to.ddTempValue
                }
            Else
                {
                $obj | Add-Member -MemberType NoteProperty -Name ConfigName$num -Value $to.ddTempName
                $obj | Add-Member -MemberType NoteProperty -Name ConfigState$num -Value $to.ddTempState
                $obj | Add-Member -MemberType NoteProperty -Name ConfigValue$num -Value $to.ddTempValue
                }
            If ($num -gt $ConfigMax)
                {
                $rsop | ForEach-Object {$_ | Add-Member -MemberType NoteProperty -Name ConfigName$num -Value $null; $_ | Add-Member -MemberType NoteProperty -Name ConfigState$num -Value $null; $_ | Add-Member -MemberType NoteProperty -Name ConfigValue$num -Value $null}
                $ConfigMax = $num
                }
            }
        Remove-Variable -Name ddTempArray
        }
    Else {}

    $rsop += $obj # adds objec to master list, then removes all temporary variables to prepare for next cycle.
    Remove-Variable -name obj
    Remove-Variable -name p
    Remove-Variable -Name GPMatch
    Remove-Variable -Name PolGPID
    }

#################################################################
## Coverts object array to CSV format, then pastes into Excel. ##
## Formats Excel table for readability.                        ##
#################################################################

$rsop | ConvertTo-Csv -Delimiter "`t" -NoTypeInformation | Clip
$cells = $as.Cells
$cells.Item(1).PasteSpecial() | Out-Null
Format-XL -xl $xl -wb $wb -as $as | Out-Null
$cells.Item(1,1).ColumnWidth = 16
$cells.Item(1,2).ColumnWidth = 25
$cells.Item(1,3).ColumnWidth = 30
$cells.Item(1,4).ColumnWidth = 25
$cells.Item(1,5).ColumnWidth = 65
$cells.Item(1,7).ColumnWidth = 135
$cells.Range('G:ZZ').ColumnWidth = 30
$Cells.EntireRow.AutoFit() | Out-Null
$range = $as.UsedRange
$range.HorizontalAlignment = -4131
$range.VerticalAlignment = -4160
$as.Name = "Comp+User Policies"
$wb.WorkSheets.Add() | Out-Null
$as = $wb.WorkSheets.Item(1)

　
#######################
## Security Settings ##
#######################

# Gets Security Settings from XML file.
$settingGroups = $xml.DocumentElement.ComputerResults.ExtensionData | Select -ExpandProperty extension | Where-Object {$_.type -like "*SecuritySettings"} | Get-Member -MemberType Property | Where-Object {$_.Name -ne "q10" -and $_.Name -ne "type" -and $_.Name -ne "xmlns"}
$results = ForEach ($sg in $settingGroups) {$xml.DocumentElement.ComputerResults.ExtensionData | Select -ExpandProperty extension | Where-Object {$_.type -like "*SecuritySettings"} | Select $sg.Name | Select -ExpandProperty *}

$propMax = 0
$resultList = @()

ForEach ($r in $results)
{
$propNum = ($r | Get-Member -MemberType Property | Measure-Object).count
If ($propNum -gt $propMax) {$propMax = $propNum}
}

ForEach ($r in $results)
{ # Creates temporary objec to hold data for current setting being parsed.
$num = 3
$obj = New-Object PSObject
$obj | Add-Member -MemberType NoteProperty -Name GPO -Value $null
$obj | Add-Member -MemberType NoteProperty -Name GPODomain -Value $null
$obj | Add-Member -MemberType NoteProperty -Name Precedence -Value ($r | Select -ExpandProperty Precedence | Select '#text').'#text'
$obj | Add-Member -MemberType NoteProperty -Name Name -Value $null

# Gets GPO name and domain for winning GPO.
$PolGPID = $r | Select GPO | Select -ExpandProperty * | Select Identifier | Select -ExpandProperty * | Select `#text
$GPMatch = $GPList | Where-Object -Property GPID -eq -Value $PolGPID.'#text'
$obj.GPO = $GPMatch.GPName
$obj.GPODomain = $GPMatch.GPDomain

$itemNum = 0
While ($num -lt $propMax) # Adds additional properties to equal number of properties in largest object.
    {
    $itemNum++
    $obj | Add-Member -MemberType NoteProperty -Name Prop$itemNum -Value $null
    $num++
    }
$propNames = $r | Get-Member -MemberType Property | Select Name
$newNum = 1

ForEach ($pn in $propNames)
    { # Parses each property for current object.
    $pn = $pn.Name
    If ($pn -like "*Name")
        {
        If ($pn -eq 'Name'-or $pn -eq 'KeyName' -or $pn -eq "SystemAccessPolicyName") {$obj.Name = $r.$pn; Continue}
        Else {$obj.Name = ($r | Select $pn | Select -ExpandProperty * | Select Name | Select -ExpandProperty * | Select '#text').'#text'; Continue}
        }
    ElseIf ($pn -eq "SecurityDescriptor")
        {
        $perPresent = ($r | Select $pn | Select -ExpandProperty * | Select PermissionsPresent | Select -ExpandProperty * | Select '#text').'#text'
        If ($perPresent -eq "false") 
            {$obj."Prop$newNum" = 'PermissionsPresent='+$perPresent; $newNum++; Continue}
        If ($perPresent -eq "true") 
            {
            $trustArray = "Permissions=`n"
            $trustees = @($r | Select $pn | Select -ExpandProperty * | Select Permissions | Select -ExpandProperty * | Select TrusteePermissions | Select -ExpandProperty * | Select Trustee | Select -ExpandProperty * | Select Name | Select -ExpandProperty * | Select '#text').'#text'
            $permType = @($r | Select $pn | Select -ExpandProperty * | Select Permissions | Select -ExpandProperty * | Select TrusteePermissions | Select -ExpandProperty * | Select Type | Select -ExpandProperty * | Select PermissionType | Select -ExpandProperty *)
            $permAccess = @($r | Select $pn | Select -ExpandProperty * | Select Permissions | Select -ExpandProperty * | Select TrusteePermissions | Select -ExpandProperty * | Select Standard | Select -ExpandProperty * | Select ServicesGroupedAccessEnum | Select -ExpandProperty *)
            $trustCount = $trustees.Count
            $tcount = 0
            While ($tcount -lt $trustCount)
                {
                $tname = $trustees[$tcount]
                $ttype = $permType[$tcount]
                $tacc = $permAccess[$tcount]
                $tcount++
                $trustArray += "Name=`"$tname`", PermType=`"$ttype`", Access=`"$tacc`""
                If ($tcount -lt $trustCount){$trustArray += "`n"}
                }
            $obj."Prop$newNum" = $trustArray
            $newNum++
            Continue
            }
        }
    ElseIf ($pn -like "Display*")
        {
        $props = ($r | Select $pn | Select -ExpandProperty *) | Get-Member -MemberType Property
        $obj."Prop$newNum" = $pn+'='+($r | Select $pn | Select -ExpandProperty * | Select "Name" | Select -ExpandProperty *)        
        $newNum++
        Continue
        }
    ElseIf ($pn -eq "Member")
        {
        $members = $r | Select Member | Select -ExpandProperty * | Select Name | Select -ExpandProperty * | Select '#text' | Select -ExpandProperty *
        $memlist = "Members=`n"
        $memcount = $members.Count
        $mcount = 0
        ForEach ($m in $members)
            {
            $mcount++
            $memlist += "$m"
            If ($mcount -lt $memcount) {$memlist += "`n"}
            }
        $obj."Prop$newNum" = "$memlist"
        $newNum++
        Continue
        }
    ElseIf ($pn -eq "SettingStrings")
    {$obj."Prop$newNum" = $pn+'='+($r | Select SettingStrings | Select -ExpandProperty * | Select Value | Select -ExpandProperty *); $newNum++; Continue}
    ElseIf ($pn -eq "SuccessAttempts")
        {$obj."Prop$newNum" = $pn+'='+$r.$pn; $newNum++; Continue}
    ElseIf ($pn -eq "FailureAttempts")
        {$obj."Prop$newNum" = $pn+'='+$r.$pn; $newNum++; Continue}
    ElseIf ($pn -ne "GPO" -and $pn -ne "Precedence") {$obj."Prop$newNum" = $pn+'='+$r.$pn; $newNum++}
    If ($error[0] -ne $null)
    {$r; $error.clear() ; Exit}
    }
$resultList += $obj
Remove-Variable -name obj
}

#################################################################
## Coverts object array to CSV format, then pastes into Excel. ##
## Formats Excel table for readability.                        ##
#################################################################
$resultList | ConvertTo-Csv -Delimiter "`t" -NoTypeInformation | Clip
$cells = $as.Cells
$cells.Item(1).PasteSpecial() | Out-Null
Format-XL -xl $xl -wb $wb -as $as | Out-Null
$cells.Item(1,1).ColumnWidth = 25
$cells.Item(1,2).ColumnWidth = 25
$cells.Item(1,4).ColumnWidth = 25
$cells.Item(1,5).ColumnWidth = 75
$cells.Item(1,6).ColumnWidth = 25
$Cells.EntireRow.AutoFit() | Out-Null
$range = $as.UsedRange
$range.HorizontalAlignment = -4131
$range.VerticalAlignment = -4160
$as.Name = "SecuritySettings"

　
###############
## Save File ##
###############
$FileExist = Test-Path -Path "$PSScriptRoot\$compName-RSOP.xlsx"
If ($FileExist -eq $true) {Remove-Item -Path "$PSScriptRoot\$compName-RSOP.xlsx"}
$xl.ActiveWorkbook.SaveAs("$PSScriptRoot\$compName-RSOP.xlsx",$xlFixedFormat) | Out-Null
$FileExist = Test-Path -Path "$PSScriptRoot\$compName-RSOP.xlsx"
If ($FileExist -eq $true) {Remove-Item -Path "$PSScriptRoot\$compName-RSOP.xml"}
$xl.Quit() | Out-Null 
