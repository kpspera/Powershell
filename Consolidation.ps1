Add-Type -AssemblyName Microsoft.Office.Interop.Excel
$xlFixedFormat = [Microsoft.Office.Interop.Excel.XlFileFormat]::xlWorkbookDefault
$xl = new-object -comobject excel.application
$xl.visible = $true
$wb = $xl.Workbooks.Add()
$files = Get-ChildItem -Path $PSScriptRoot\* -Include *.csv

ForEach ($f in $files)
    {
    $wb.WorkSheets.Add() | Out-Null
    $as = $wb.WorkSheets.Item(1)
    if ($as.Name -eq "Sheet2") {$wb.WorkSheets.Item("Sheet1").Delete()}
    $cells = $as.Cells
    $input = Get-Content -Path $f
    $input | Clip
    $cells.Item(1).PasteSpecial() | Out-Null
    $range = $as.UsedRange
    $table = $as.ListObjects.add(1,$range,0,1)
    $range.EntireColumn.Autofit() | Out-Null
    $SheetName = $f.Name
    $as.Name = $SheetName -replace '.csv',''
    $input = $null
    }
$xl.ActiveWorkbook.SaveAs("$PSScriptRoot\$env:COMPUTERNAME-ConsolidatedSCriptResults.xlsx",$xlFixedFormat) | Out-Null
$xl.Quit() | Out-Null
