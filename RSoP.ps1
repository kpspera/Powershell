######################################################################################
##                        RESULTANT SET of POLICY XML PARSER                        ##
##                                by Ken Spera, CISA                                ##
##                                                                                  ##
## Last Update: 6/23/2017                                                           ##
##                                                                                  ##
## Description:                                                                     ##
## Parses XML-formatted Resutant Set of Policy reports (GPResult) into more useful  ##	
## Excel (.xlsx) format. For each policy returned in the RSoP Report, the policy    ##
## type (Computer vs. User), winning GPO, policy name, state, description,          ##
## supported versions, and all configurations falling under the policy are returned ##
## and formatted to an Excel table.                                                 ##
######################################################################################



Add-Type -AssemblyName Microsoft.Office.Interop.Excel
$xlFixedFormat = [Microsoft.Office.Interop.Excel.XlFileFormat]::xlWorkbookDefault
$xl = new-object -comobject excel.application
$rsop = @()

　
$xml = [xml](Get-Content $env:USERPROFILE\Desktop\RSOP.xml)
$ConfigMax = 0

$GPOs = $xml.DocumentElement.ComputerResults.GPO
$GPList = @()
ForEach ($gp in $GPOs)
    {
    $obj = New-Object PSObject
    $obj | Add-Member -MemberType NoteProperty -Name GPName -Value $gp.Name
    $obj | Add-Member -MemberType NoteProperty -Name GPEnabled -Value $gp.Enabled
    $obj | Add-Member -MemberType NoteProperty -Name GPIsValid -Value $gp.IsValid
    $obj | Add-Member -MemberType NoteProperty -Name GPID -Value $null
    $ObjID = $gp | Select Path | Select -ExpandProperty * | Select Identifier | Select -ExpandProperty * | Select `#text
    $obj.GPID = $objID.'#text'
    $GPList += $obj
    Remove-Variable -Name obj
    }

#############################
## PARSE COMPUTER SETTINGS ##
#############################
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
    $obj | Add-Member -MemberType NoteProperty -Name Category -Value $p.Category
    $obj | Add-Member -MemberType NoteProperty -Name Name -Value $p.Name
    $obj | Add-Member -MemberType NoteProperty -Name State -Value $p.State
    $obj | Add-Member -MemberType NoteProperty -Name Explain -Value $p.Explain
    $obj | Add-Member -MemberType NoteProperty -Name Supported -Value $p.Supported

    $PolGPID = $p | Select GPO | Select -ExpandProperty * | Select Identifier | Select -ExpandProperty * | Select `#text
    $GPMatch = $GPList | Where-Object -Property GPID -eq -Value $PolGPID.'#text'
    $obj.WinningGPO = $GPMatch.GPName

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

    $rsop += $obj
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
    $obj | Add-Member -MemberType NoteProperty -Name Category -Value $p.Category
    $obj | Add-Member -MemberType NoteProperty -Name Name -Value $p.Name
    $obj | Add-Member -MemberType NoteProperty -Name State -Value $p.State
    $obj | Add-Member -MemberType NoteProperty -Name Explain -Value $p.Explain
    $obj | Add-Member -MemberType NoteProperty -Name Supported -Value $p.Supported

    $PolGPID = $p | Select GPO | Select -ExpandProperty * | Select Identifier | Select -ExpandProperty * | Select `#text
    $GPMatch = $GPList | Where-Object -Property GPID -eq -Value $PolGPID.'#text'
    $obj.WinningGPO = $GPMatch.GPName

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

    $rsop += $obj
    Remove-Variable -name obj
    Remove-Variable -name p
    Remove-Variable -Name GPMatch
    Remove-Variable -Name PolGPID
    }

$rsop | ConvertTo-Csv -Delimiter "`t" -NoTypeInformation | Clip
$xl.visible = $true
$wb = $xl.Workbooks.Add()
$as = $wb.WorkSheets.Item(1)
$cells = $as.Cells
$cells.Item(1).PasteSpecial() | Out-Null
$range = $as.UsedRange
$range.WrapText = $true
$cells.EntireColumn.AutoFit() | Out-Null
$cells.Item(1,1).ColumnWidth = 16
$cells.Item(1,2).ColumnWidth = 25
$cells.Item(1,3).ColumnWidth = 25
$cells.Item(1,4).ColumnWidth = 65
$cells.Item(1,6).ColumnWidth = 135
$cells.Range('G:ZZ').ColumnWidth = 30
$Cells.EntireRow.AutoFit() | Out-Null
$range.VerticalAlignment = -4160
$atable = $as.ListObjects.add(1,$range,0,1)
switch ($xl.Version)
    {
    '15.0' {$atable.TableStyle = "TableStyleMedium21"}
    '14.0' {$atable.TableStyle = "TableStyleMedium18"}
    }
$FileExist = Test-Path -Path $env:USERPROFILE\Desktop\RSoP-to-Excel.xlsx
If ($FileExist -eq $true) {Remove-Item -Path $env:USERPROFILE\Desktop\RSoP-to-Excel.xlsx}
$xl.ActiveWorkbook.SaveAs("$env:USERPROFILE\Desktop\RSoP-to-Excel.xlsx",$xlFixedFormat) | Out-Null
$xl.Quit() | Out-Null 
