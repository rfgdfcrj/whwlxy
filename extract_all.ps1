Add-Type -AssemblyName System.IO.Compression.FileSystem

function Extract-DocxText {
    param([string]$docxPath, [string]$outputDir)
    
    $zipPath = Join-Path $env:TEMP "temp_extract.zip"
    Copy-Item $docxPath $zipPath -Force
    
    if (Test-Path $outputDir) { Remove-Item $outputDir -Recurse -Force }
    [System.IO.Compression.ZipFile]::ExtractToDirectory($zipPath, $outputDir)
    Remove-Item $zipPath -Force
    
    $xmlPath = Join-Path $outputDir "word\document.xml"
    [xml]$xml = Get-Content $xmlPath -Encoding UTF8
    
    $ns = @{w = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"}
    $texts = $xml.SelectNodes("//w:t", $ns) | ForEach-Object { $_.InnerText }
    
    return ($texts -join "")
}

$baseDir = "D:\解压文件夹\2026招生宣传资料"
$docsDir = Join-Path $baseDir "学院简介+专业介绍"

# Get all docx files
$files = Get-ChildItem -Path $docsDir -Filter "*.docx"
$outFile = Join-Path $baseDir "专业介绍_提取.txt"

$results = @()

foreach ($file in $files) {
    $extractDir = Join-Path $baseDir ("temp_" + [System.IO.Path]::GetRandomFileName())
    
    Write-Host "Extracting: $($file.Name)..."
    $text = Extract-DocxText -docxPath $file.FullName -outputDir $extractDir
    
    $results += "===== $($file.BaseName) ====="
    $results += $text
    $results += "`n"
}

$results -join "`n" | Out-File -FilePath $outFile -Encoding UTF8
Write-Host "Done! Output: $outFile"
