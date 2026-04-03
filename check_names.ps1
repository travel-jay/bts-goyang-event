[Console]::OutputEncoding = [System.Text.Encoding]::UTF8
$OutputEncoding = [System.Text.Encoding]::UTF8

$d = Get-Content "data_new_raw.json" -Raw -Encoding UTF8 | ConvertFrom-Json

$zhSameAsEn = 0
$jaSameAsEn = 0
$diff = @()

foreach ($row in $d) {
    $isSame = ($row.nameZh -eq $row.nameEn) -and ($row.nameJa -eq $row.nameEn)
    if ($row.nameZh -ne $row.nameEn -or $row.nameJa -ne $row.nameEn) {
        $diff += "No$($row.no): KO=$($row.nameKo) | EN=$($row.nameEn) | ZH=$($row.nameZh) | JA=$($row.nameJa)"
    }
    if ($row.nameZh -eq $row.nameEn) { $zhSameAsEn++ }
    if ($row.nameJa -eq $row.nameEn) { $jaSameAsEn++ }
}

Write-Host "ZH same as EN: $zhSameAsEn / $($d.Count)"
Write-Host "JA same as EN: $jaSameAsEn / $($d.Count)"
Write-Host "Stores with different ZH or JA name: $($diff.Count)"
foreach ($x in $diff) { Write-Host $x }
