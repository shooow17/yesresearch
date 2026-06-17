$ErrorActionPreference = "Stop"
$scriptPath = $PSScriptRoot
$pdfDir = Join-Path -Path $scriptPath -ChildPath "pdf_files"

if (-not (Test-Path -Path $pdfDir)) {
    New-Item -ItemType Directory -Path $pdfDir | Out-Null
}

$word = New-Object -ComObject Word.Application
$word.Visible = $false
$word.DisplayAlerts = 0

$docxFiles = Get-ChildItem -Path $scriptPath -Filter "*.docx"
$total = $docxFiles.Count
$count = 0

Write-Host "Found $total docx files. Starting conversion..."

$tree = @()

foreach ($file in $docxFiles) {
    $count++
    $pdfFileName = $file.BaseName + ".pdf"
    $pdfPath = Join-Path -Path $pdfDir -ChildPath $pdfFileName
    
    $tree += @{
        "name" = $file.BaseName
        "pdf" = "pdf_files/$pdfFileName"
    }

    if (-not (Test-Path -Path $pdfPath)) {
        Write-Host "[$count/$total] Converting: $($file.Name)"
        try {
            $pathStr = $pdfPath.ToString()
            $format = 17
            $doc = $word.Documents.Open($file.FullName, $null, $true)
            $doc.SaveAs([ref]$pathStr, [ref]$format)
            $doc.Close([ref]0)
        } catch {
            Write-Host "Error converting $($file.Name): $_"
        }
    } else {
        Write-Host "[$count/$total] Skipping, PDF exists: $($file.Name)"
    }
}

$word.Quit()
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($word) | Out-Null

Write-Host "Conversion completed."

$jsonPath = Join-Path -Path $scriptPath -ChildPath "data.json"
$tree | ConvertTo-Json -Depth 10 | Out-File -FilePath $jsonPath -Encoding UTF8

Write-Host "data.json generated."
