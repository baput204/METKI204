# update-data.ps1
# Uruchom ten skrypt aby zaktualizowac data.js z aktualnego metki.xls
# Potem: git add data.js && git commit -m "update metki data" && git push

$xlsPath = "\\RETPL204-NT0001.ikea.com\Common\STOCKOWCY 204\#NARZEDZIA#\METKI\metki.xls"
$outPath  = Join-Path $PSScriptRoot "data.js"

Write-Host "Otwieranie $xlsPath ..."

$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$excel.DisplayAlerts = $false
$wb = $excel.Workbooks.Open($xlsPath, 0, $true)
$ws = $wb.Sheets.Item("Data")
$rows = $ws.UsedRange.Rows.Count
$cols = $ws.UsedRange.Columns.Count

$headers = @()
for ($c = 1; $c -le $cols; $c++) { $headers += $ws.Cells.Item(1, $c).Value2 }

$all = @()
for ($r = 2; $r -le $rows; $r++) {
    $obj = [ordered]@{}
    for ($c = 1; $c -le $cols; $c++) {
        $v = $ws.Cells.Item($r, $c).Value2
        $obj[$headers[$c-1]] = if ($v -eq $null) { "" } else { $v }
    }
    $all += $obj
}

$wb.Close($false)
$excel.Quit()
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null

# Filtruj puste wiersze
$all = $all | Where-Object { $_['NAZWA ART.'] -ne '' -or ($_['NR ART.'] -replace '^0+','') -ne '' }

$json = $all | ConvertTo-Json -Compress
$dateStr = Get-Date -Format "dd.MM.yyyy"
$content = "const METKI_DATA_DATE = `"$dateStr`";`nconst METKI_DATA = $json;"
# Zapisz UTF-8 BEZ BOM (Out-File dodaje BOM w PS 5.1)
[System.IO.File]::WriteAllText($outPath, $content, (New-Object System.Text.UTF8Encoding $false))

Write-Host "Zapisano $($all.Count) wierszy do $outPath"
Write-Host ""
Write-Host "Teraz wykonaj:"
Write-Host "  git add data.js"
Write-Host "  git commit -m `"update: metki $(Get-Date -Format 'yyyy-MM-dd')`""
Write-Host "  git push"
