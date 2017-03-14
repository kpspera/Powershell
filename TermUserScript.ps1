#################################################################################################
###                          Terminated User Comparison Script                                ###
###            Written By: Kenneth P. Spera, CISA         Last Update: 3/13/2017              ###
#################################################################################################

# Load assemblies and define general variables.

$error.clear() | Out-Null
Add-Type -AssemblyName Microsoft.Office.Interop.Excel
[System.Reflection.Assembly]::LoadWithPartialName("System.windows.forms") | Out-Null
$initialDirectory = $PSScriptRoot
$RunDate = Get-Date -Format "MM/dd/yyyy hh:mm tt"

# Use Open File dialog to select source HR extract file and collect file creation date.

Write-Verbose -Message "Terminated User comparison script initalizing. You will be prompted to select the source HR extract to use." -Verbose
"`n"
Pause
"`n"
$OpenDialog = New-Object System.Windows.Forms.OpenFileDialog
$dialogResult = $OpenDialog.ShowDialog()
If ($dialogResult -ne "OK")
    {
    Write-Warning "ERROR: User canceled file selection. Quitting....." -Verbose
    "`n"
    Pause
    Exit
    }
$file = $OpenDialog.FileName
$fileName = $OpenDialog.SafeFileName
$HRFile = Get-ItemProperty -Path $file
$HRDate = $HRFile.CreationTime
$HRDate = Get-Date $HRDate -Format "MM/dd/yyyy hh:mm tt"

# Prepare to launch Excel to begin comparison.

"STATUS: Opening HR file `"$fileName`"`n"
$xlFixedFormat = [Microsoft.Office.Interop.Excel.XlFileFormat]::xlWorkbookDefault
$xl = new-object -comobject excel.application
$xl.visible = $false

# Open HR file in Excel and define Excel-related variables.

$wb = $xl.workbooks.open("$file")
$ActiveSheet = $wb.ActiveSheet
$range = $ActiveSheet.UsedRange
$table = $ActiveSheet.ListObjects.add(1,$range,0,1)
$Cells = $ActiveSheet.Cells

# Import list of user IDs to compare to HR extract.

$inputs = Get-Content "$PSScriptRoot\input.txt"
$inputs = $inputs | Sort-Object
$inCount = $inputs.Count

# Find input user information in HR.

"STATUS: Getting HR data for $inCount users.`n"
$Users = @()
ForEach ($i in $inputs)
    {
    $UsrObj = New-Object PSObject
    $UsrObj | Add-Member -MemberType NoteProperty -Name SysID -Value $null
    $UsrObj | Add-Member -MemberType NoteProperty -Name HRID -Value $null
    $UsrObj | Add-Member -MemberType NoteProperty -Name HRLN -Value $null
    $UsrObj | Add-Member -MemberType NoteProperty -Name HRFN -Value $null
    $UsrObj | Add-Member -MemberType NoteProperty -Name HRES -Value $null
    $UsrObj | Add-Member -MemberType NoteProperty -Name HRTD -Value $null
    $UsrObj | Add-Member -MemberType NoteProperty -Name HRJT -Value $null
    $UsrObj | Add-Member -MemberType NoteProperty -Name HRDN -Value $null
    $i = $i.ToUpper()
    $UsrObj.SysID = $i
    $Result = $Cells.Find("$i")
    $row = $Result.Row 
    if ($row -eq $null)
        {
        Write-Verbose -message "Error finding user `"$i`". Please check that user is not a Generic, system, or service ID. Noting error and continuing with next user." -Verbose
        "`n"
        $UsrObj.HRID = '*** ERROR: User name could not be found in HR file. Possibly Generic or system/service ID. ***'
        $UsrObj.HRFN = $null
        $UsrObj.HRLN = $null
        $UsrObj.HRES = $null
        $UsrObj.HRTD = $null
        $UsrObj.HRJT = $null
        $UsrObj.HRDN = $null
        $Users += $UsrObj
        Continue
        }
    $UsrObj.HRID = $Cells.Item($row,1).Value()
    $UsrObj.HRFN = $Cells.Item($row,2).Value()
    $UsrObj.HRLN = $Cells.Item($row,4).Value()
    $UsrObj.HRES = $Cells.Item($row,5).Value()
    $UsrObj.HRTD = $Cells.Item($row,9).Value()
    $UsrObj.HRJT = $Cells.Item($row,11).Value()
    $UsrObj.HRDN = $Cells.Item($row,15).Value()
    $Users += $UsrObj
    Remove-Variable -Name "UsrObj" | Out-Null
    }
[Void]$wb.close($false)

# Convert input user HR data to CSV format, tab delimited.

$UserCount = $Users.Count
"STATUS: Loading data for $UserCount users for analysis.`n"
$Users | ConvertTo-Csv -Delimiter "`t" -NoTypeInformation | Clip

# Load CSV data into new Excel workbook and apply basic formatting.

$wb2 = $xl.Workbooks.Add()
$AS2 = $wb2.Worksheets.Item(1)
[Void]$AS2.Cells.Item(1).PasteSpecial()
$range2 = $AS2.UsedRange
$table2 = $AS2.ListObjects.add(1,$range2,0,1)
$MaxRows2 = $range2.Rows.Count
$Cells2 = $AS2.Cells
[Void]$range2.EntireColumn.Autofit()

# Cyle through all users, determine if user is terminated, if so, highlight in red with white text.

"STATUS: Analysing $UserCount users.`n"
$row2 = 1
while ($row2 -lt $MaxRows2)
    {
    $row2++
    If ($Cells2.Item($row2,5).Value() -ne $null -and $Cells2.Item($row2,5).Value() -ne "T") {Continue}
    ElseIf ($Cells2.Item($row2,5).Value() -eq $null)
        {
        $TermRange = $xl.Range("A${row2}:H$row2")
        $TermRange.Interior.ColorIndex = 6
        }
    Else
        {
        $TermRange = $xl.Range("A${row2}:H$row2")
        $TermRange.Interior.ColorIndex = 3
        $TermRange.Font.ColorIndex = 2
        }
    }

# Insert legend and source information at top of sheet.

"STATUS: Adding legend and source information to output file.`n"
$xlShiftDown = -4121
$eRow = $Cells2.item(1,1).entireRow
$eRow.insert($xlShiftDown) | Out-Null
$eRow.insert($xlShiftDown) | Out-Null
For ($i=1;$i -le 3;$i++)
    {
    $eRow = $Cells2.item(1,1).entireRow
    $eRow.insert($xlShiftDown) | Out-Null
    $Cells2.item(1,1).BorderAround(1,2,1) | Out-Null
    $Cells2.item(1,2).BorderAround(1,2,1) | Out-Null
    }
$Cells2.item(1,1) = "Legend"
$Cells2.item(1,2) = "Color Code"
$Cells2.item(2,1) = "Terminated User"
$Cells2.item(2,2) = "<Text>"
$Cells2.item(3,1) = "Error"
$Cells2.item(3,2) = "<Text>"
$Cells2.item(4,1) = "Total Users: $UserCount"
$Cells2.item(1,5) = "Source File: $fileName"
$Cells2.item(2,5) = "Source File Date: $HRDate"
$Cells2.item(3,5) = "Script Run Date: $RunDate"
$ActiveRange = $Cells2.Range("A1","B1")
$ActiveRange.Interior.ColorIndex = 10
$Cells2.Range("A1","B1").Font.ColorIndex = 2
$Cells2.Range("A1","D1").EntireColumn.AutoFit() | Out-Null
$Cells2.item(1,2).ColumnWidth = 15
$Cells2.item(2,2).Interior.ColorIndex = 3
$Cells2.item(2,2).Font.ColorIndex = 2
$Cells2.item(3,2).Interior.ColorIndex = 6

# Save HR comparison file and close Excel.

$xl.ActiveWorkbook.SaveAs("$PSScriptRoot\HRCompare.xlsx",$xlFixedFormat)
[Void]$wb2.close($true)
[Void]$xl.Quit()
Write-Verbose -Message "Analysis complete. See file `"$PSScriptRoot\HRCompare.xlsx`" for results." -Verbose

# Clean up any orphaned Excel processes.

[System.Runtime.Interopservices.Marshal]::ReleaseComObject($xl) | Out-Null
Stop-Process -Name EXCEL |Out-Null
Remove-Variable xl | Out-Null
[System.GC]::Collect() | Out-Null
"`n"
Pause
Exit
