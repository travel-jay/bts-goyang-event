[Console]::OutputEncoding = [System.Text.Encoding]::UTF8
$json = Get-Content "data_merged.json" -Raw -Encoding UTF8
# Already compressed from ConvertTo-Json -Compress, just save without BOM
$utf8NoBom = New-Object System.Text.UTF8Encoding $false
[System.IO.File]::WriteAllText("C:\Users\groun\claude_code_projects\26.04_test_project\data_min.json", $json, $utf8NoBom)
Write-Host "Saved data_min.json, size=$(((Get-Item 'data_min.json').Length)) bytes"
