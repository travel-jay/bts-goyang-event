[Console]::OutputEncoding = [System.Text.Encoding]::UTF8
$OutputEncoding = [System.Text.Encoding]::UTF8

$newData = Get-Content "data_new_raw.json" -Raw -Encoding UTF8 | ConvertFrom-Json
$oldData = Get-Content "data_geocoded.json" -Raw -Encoding UTF8 | ConvertFrom-Json

# Build lookup by no
$oldMap = @{}
foreach ($row in $oldData) { $oldMap[$row.no] = $row }

$merged = @()
foreach ($row in $newData) {
    $old = $oldMap[$row.no]
    $lat = if ($old) { $old.lat } else { $null }
    $lng = if ($old) { $old.lng } else { $null }

    # Use existing nameEn (has bilingual pattern like "영덕해물전골 (Yeongdeokhaemuljeongol)")
    # If old nameEn differs from new (has bilingual), keep old pattern
    $nameEn = if ($old -and $old.nameEn) { $old.nameEn } else { $row.nameEn }

    # For ZH and JA: nameZh/nameJa are same as EN in the Excel
    # Apply same bilingual rule: if nameEn contains "(", nameZh/nameJa should too
    # Use same nameEn pattern for ZH/JA since Excel has no distinct ZH/JA names
    $nameZh = if ($row.nameZh -and $row.nameZh -ne $row.nameEn) { $row.nameZh } else { $nameEn }
    $nameJa = if ($row.nameJa -and $row.nameJa -ne $row.nameEn) { $row.nameJa } else { $nameEn }

    $item = [ordered]@{
        no        = $row.no
        nameKo    = $row.nameKo
        nameEn    = $nameEn
        nameZh    = $nameZh
        nameJa    = $nameJa
        typeKo    = $row.typeKo
        typeEn    = if ($old -and $old.typeEn) { $old.typeEn } else { $row.typeEn }
        typeZh    = $row.typeZh
        typeJa    = $row.typeJa
        addressKo = $row.addressKo
        addressEn = if ($old -and $old.addressEn) { $old.addressEn } else { $row.addressEn }
        addressZh = $row.addressZh
        addressJa = $row.addressJa
        eventKo   = $row.eventKo
        eventEn   = if ($old -and $old.eventEn) { $old.eventEn } else { $row.eventEn }
        eventZh   = $row.eventZh
        eventJa   = $row.eventJa
        lat       = $lat
        lng       = $lng
    }
    $merged += $item
}

$missingCoords = ($merged | Where-Object { $_.lat -eq $null }).Count
Write-Host "Total: $($merged.Count), Missing coords: $missingCoords"

# Sample check
Write-Host "Sample[2]: nameKo=$($merged[2].nameKo) nameEn=$($merged[2].nameEn) typeZh=$($merged[2].typeZh) lat=$($merged[2].lat)"

$json = $merged | ConvertTo-Json -Depth 3 -Compress
[System.IO.File]::WriteAllText("C:\Users\groun\claude_code_projects\26.04_test_project\data_merged.json", $json, [System.Text.Encoding]::UTF8)
Write-Host "Saved data_merged.json"
