$xlsxPath = "d:\Monash\FIT2179\data\chart 11 - dot density map.xlsx"
$csvPath  = "d:\Monash\FIT2179\data\chart11_dot_density_points.csv"

Write-Host "Opening workbook $xlsxPath..."
$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$excel.DisplayAlerts = $false

$wb = $excel.Workbooks.Open($xlsxPath)
$sheet = $wb.Sheets.Item("Dot_Points")

Write-Host "Reading geocoded dot points..."
$results = @()

$lastRow = $sheet.UsedRange.Rows.Count
Write-Host "Found $lastRow rows in Dot_Points."

# Headers are: dot_id, longitude, latitude, SA2_Code, SA2_Name, State_Name, NOM_2024_25, Dot_Value
for ($r = 2; $r -le $lastRow; $r++) {
    $dotId      = $sheet.Cells.Item($r, 1).Text
    $longitude  = $sheet.Cells.Item($r, 2).Text
    $latitude   = $sheet.Cells.Item($r, 3).Text
    $sa2Code    = $sheet.Cells.Item($r, 4).Text
    $sa2Name    = $sheet.Cells.Item($r, 5).Text
    $stateName  = $sheet.Cells.Item($r, 6).Text
    $nom        = $sheet.Cells.Item($r, 7).Text
    $dotValue   = $sheet.Cells.Item($r, 8).Text
    
    if ($dotId -and $longitude -and $latitude) {
        $results += [PSCustomObject]@{
            dot_id       = $dotId
            longitude    = [double]$longitude
            latitude     = [double]$latitude
            SA2_Code     = $sa2Code
            SA2_Name     = $sa2Name
            State_Name   = $stateName
            NOM_2024_25  = [int]$nom
            Dot_Value    = [int]$dotValue
        }
    }
}

$wb.Close($false)
$excel.Quit()
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null

Write-Host "Extracted $($results.Count) dot points. Exporting to CSV..."
$results | Export-Csv -Path $csvPath -NoTypeInformation -Encoding utf8
Write-Host "Success! Geocoded dot density coordinates exported to $csvPath."
