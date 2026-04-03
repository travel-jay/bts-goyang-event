[Console]::OutputEncoding = [System.Text.Encoding]::UTF8
$OutputEncoding = [System.Text.Encoding]::UTF8

$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$excel.DisplayAlerts = $false

$wb = $excel.Workbooks.Open("C:\Users\groun\claude_code_projects\26.04_test_project\new_data.xlsx")

$sheets = @{
    ko = $wb.Sheets.Item(1)
    en = $wb.Sheets.Item(2)
    zh = $wb.Sheets.Item(3)
    ja = $wb.Sheets.Item(4)
}

$lastRow = $sheets.ko.UsedRange.Rows.Count
$rows = @()

for ($r = 2; $r -le $lastRow; $r++) {
    $no      = [int]$sheets.ko.Cells.Item($r, 1).Value2
    $nameKo  = $sheets.ko.Cells.Item($r, 2).Value2
    $typeKo  = $sheets.ko.Cells.Item($r, 3).Value2
    $addrKo  = $sheets.ko.Cells.Item($r, 4).Value2
    $eventKo = $sheets.ko.Cells.Item($r, 5).Value2

    $nameEn  = $sheets.en.Cells.Item($r, 2).Value2
    $typeEn  = $sheets.en.Cells.Item($r, 3).Value2
    $addrEn  = $sheets.en.Cells.Item($r, 4).Value2
    $eventEn = $sheets.en.Cells.Item($r, 5).Value2

    $nameZh  = $sheets.zh.Cells.Item($r, 2).Value2
    $typeZh  = $sheets.zh.Cells.Item($r, 3).Value2
    $addrZh  = $sheets.zh.Cells.Item($r, 4).Value2
    $eventZh = $sheets.zh.Cells.Item($r, 5).Value2

    $nameJa  = $sheets.ja.Cells.Item($r, 2).Value2
    $typeJa  = $sheets.ja.Cells.Item($r, 3).Value2
    $addrJa  = $sheets.ja.Cells.Item($r, 4).Value2
    $eventJa = $sheets.ja.Cells.Item($r, 5).Value2

    $row = [ordered]@{
        no      = $no
        nameKo  = $nameKo
        nameEn  = $nameEn
        nameZh  = $nameZh
        nameJa  = $nameJa
        typeKo  = $typeKo
        typeEn  = $typeEn
        typeZh  = $typeZh
        typeJa  = $typeJa
        addressKo = $addrKo
        addressEn = $addrEn
        addressZh = $addrZh
        addressJa = $addrJa
        eventKo = $eventKo
        eventEn = $eventEn
        eventZh = $eventZh
        eventJa = $eventJa
    }
    $rows += $row
}

$json = $rows | ConvertTo-Json -Depth 3
[System.IO.File]::WriteAllText("C:\Users\groun\claude_code_projects\26.04_test_project\data_new_raw.json", $json, [System.Text.Encoding]::UTF8)
Write-Host "Done. Total rows: $($rows.Count)"

$wb.Close($false)
$excel.Quit()
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null
