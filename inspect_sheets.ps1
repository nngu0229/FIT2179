$xlsxPath = "d:\Monash\FIT2179\data\chart 11 - dot density map.xlsx"
$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$wb = $excel.Workbooks.Open($xlsxPath)
foreach ($sheet in $wb.Sheets) {
    Write-Host $sheet.Name
}
$wb.Close($false)
$excel.Quit()
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null
