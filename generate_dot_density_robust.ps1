$postcodePath = "d:\Monash\FIT2179\data\temp_postcodes.csv"
$xlsxPath     = "d:\Monash\FIT2179\data\chart 11 - dot density map.xlsx"
$outputPath   = "d:\Monash\FIT2179\data\chart11_dot_density_points.csv"

Write-Host "Step 1: Reading geocoded postcodes from CSV to calculate SA2 centroids..."
$sa2LatSum = @{}
$sa2LonSum = @{}
$sa2Count  = @{}

# Read CSV and build SA2 mapping in a clean, scalar way
$csv = Import-Csv -Path $postcodePath
foreach ($row in $csv) {
    $sa2Code = $row.SA2_CODE_2021
    $lat = $row.Lat_precise
    $lon = $row.Long_precise
    
    if ($sa2Code -and $sa2Code.Length -eq 9 -and $lat -and $lon) {
        $latD = [double]$lat
        $lonD = [double]$lon
        
        if (-not $sa2LatSum.Contains($sa2Code)) {
            $sa2LatSum[$sa2Code] = 0.0
            $sa2LonSum[$sa2Code] = 0.0
            $sa2Count[$sa2Code] = 0
        }
        
        $sa2LatSum[$sa2Code] += $latD
        $sa2LonSum[$sa2Code] += $lonD
        $sa2Count[$sa2Code]  += 1
    }
}

Write-Host "Calculating mean centroids for each SA2..."
$sa2Centroids = @{}
foreach ($key in $sa2LatSum.Keys) {
    $meanLat = $sa2LatSum[$key] / $sa2Count[$key]
    $meanLon = $sa2LonSum[$key] / $sa2Count[$key]
    $sa2Centroids[$key] = @($meanLat, $meanLon)
}

Write-Host "Step 2: Loading cleaned migration data from Excel ($xlsxPath)..."
$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$excel.DisplayAlerts = $false
$wb = $excel.Workbooks.Open($xlsxPath)
$sheet = $wb.Sheets.Item(1)

$migrationData = @()
# Find total rows
$lastRow = $sheet.UsedRange.Rows.Count
Write-Host "Found $lastRow rows in Excel."

for ($r = 2; $r -le $lastRow; $r++) {
    $stateName = $sheet.Cells.Item($r, 1).Text
    $gccsaName = $sheet.Cells.Item($r, 2).Text
    $sa2Code   = $sheet.Cells.Item($r, 3).Text
    $sa2Name   = $sheet.Cells.Item($r, 4).Text
    
    $arrivals   = $sheet.Cells.Item($r, 5).Value2
    $departures = $sheet.Cells.Item($r, 6).Value2
    $nom        = $sheet.Cells.Item($r, 7).Value2
    
    # Strip any text prefix quote if present
    if ($sa2Code -and $sa2Code.StartsWith("'")) {
        $sa2Code = $sa2Code.Substring(1)
    }
    
    if ($sa2Code -and $sa2Code.Length -eq 9) {
        $migrationData += [PSCustomObject]@{
            State_Name = $stateName
            GCCSA_Name = $gccsaName
            SA2_Code   = $sa2Code
            SA2_Name   = $sa2Name
            Arrivals   = if ($arrivals -ne $null) { [int]$arrivals } else { 0 }
            Departures = if ($departures -ne $null) { [int]$departures } else { 0 }
            NOM        = if ($nom -ne $null) { [int]$nom } else { 0 }
        }
    }
}

$wb.Close($false)
$excel.Quit()
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null

Write-Host "Step 3: Generating random dots based on migration arrivals (1 dot = 100 arrivals)..."
$dots = @()
$dotValue = 100 # 1 dot represents 100 arrivals

$rand = New-Object System.Random

foreach ($row in $migrationData) {
    $sa2Code = $row.SA2_Code
    
    if ($sa2Centroids.Contains($sa2Code)) {
        $centroid = $sa2Centroids[$sa2Code]
        $lat = $centroid[0]
        $lon = $centroid[1]
        
        $arrivals = $row.Arrivals
        $numDots = [int]($arrivals / $dotValue)
        
        # Ensure that regions with high arrivals get points, and smaller regions get at least 1 dot if they have arrivals
        if ($arrivals -gt 0 -and $numDots -eq 0) {
            if ($rand.Next(0, 100) -lt ($arrivals / $dotValue * 100)) {
                $numDots = 1
            }
        }
        
        # Generate random spatial dispersion around the centroid
        for ($i = 0; $i -lt $numDots; $i++) {
            # Dispersion radius
            $dispersion = 0.015 # Capital City (about 1.5 km radius)
            if ($row.GCCSA_Name -match "Rest of") {
                $dispersion = 0.04 # Regional (about 4 km radius)
            }
            
            $offsetLat = ($rand.NextDouble() - 0.5) * 2 * $dispersion
            $offsetLon = ($rand.NextDouble() - 0.5) * 2 * $dispersion
            
            $dotLat = $lat + $offsetLat
            $dotLon = $lon + $offsetLon
            
            $dots += [PSCustomObject]@{
                State_Name = $row.State_Name
                GCCSA_Name = $row.GCCSA_Name
                SA2_Code   = $row.SA2_Code
                SA2_Name   = $row.SA2_Name
                Lat        = [Math]::Round($dotLat, 5)
                Lon        = [Math]::Round($dotLon, 5)
                Arrivals   = $row.Arrivals
                NOM        = $row.NOM
            }
        }
    }
}

Write-Host "Generated $($dots.Count) dot density coordinates. Exporting to CSV..."
$dots | Export-Csv -Path $outputPath -NoTypeInformation -Encoding utf8
Write-Host "Success! Geocoded dot density coordinate file saved to $outputPath."
