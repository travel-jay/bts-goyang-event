[Console]::OutputEncoding = [System.Text.Encoding]::UTF8
$OutputEncoding = [System.Text.Encoding]::UTF8

$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$excel.DisplayAlerts = $false

$wb = $excel.Workbooks.Open("C:\Users\groun\claude_code_projects\26.04_test_project\new_data.xlsx")

# List all sheets
$sheetCount = $wb.Sheets.Count
Write-Host "SheetCount=$sheetCount"
for ($s = 1; $s -le $sheetCount; $s++) {
    Write-Host "Sheet${s}=$($wb.Sheets.Item($s).Name)"
}

# Read all sheets info
for ($s = 1; $s -le $sheetCount; $s++) {
    $ws = $wb.Sheets.Item($s)
    $lastRow = $ws.UsedRange.Rows.Count
    $lastCol = $ws.UsedRange.Columns.Count
    Write-Host "=Sheet$s($($ws.Name)) Rows=$lastRow Cols=$lastCol"
    for ($c = 1; $c -le $lastCol; $c++) {
        $val = $ws.Cells.Item(1, $c).Value2
        Write-Host "  H${c}=$val"
    }
    # First 2 data rows
    for ($r = 2; $r -le [Math]::Min(3, $lastRow); $r++) {
        Write-Host "  Row${r}:"
        for ($c = 1; $c -le $lastCol; $c++) {
            $val = $ws.Cells.Item($r, $c).Value2
            Write-Host "    F${c}=$val"
        }
    }
}

$wb.Close($false)
$excel.Quit()
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null
