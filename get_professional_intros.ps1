Add-Type -AssemblyName System.IO.Compression.FileSystem

$docsDir = "D:\解压文件夹\2026招生宣传资料\学院简介+专业介绍"
$outFile = "D:\解压文件夹\2026招生宣传资料\all_professional_intros.txt"

$output = @()

# List all docx files
$files = Get-ChildItem -Path $docsDir -Filter "*.docx"

foreach ($file in $files) {
    Write-Host "Processing: $($file.Name)"
    
    $extractDir = Join-Path "D:\解压文件夹\2026招生宣传资料" "temp_extract"
    if (Test-Path $extractDir) { Remove-Item $extractDir -Recurse -Force }
    
    [System.IO.Compression.ZipFile]::ExtractToDirectory($file.FullName, $extractDir)
    
    $xmlPath = Join-Path $extractDir "word\document.xml"
    $xml = [xml](Get-Content $xmlPath -Encoding UTF8)
    
    $ns = @{w = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"}
    $texts = $xml.SelectNodes("//w:t", $ns) | ForEach-Object { $_.InnerText }
    $text = $texts -join ""
    
    $output += "===== $($file.Name) ====="
    $output += $text
    $output += "`n`n"
    
    Remove-Item $extractDir -Recurse -Force
}

$output -join "`n" | Out-File -FilePath $outFile -Encoding UTF8
Write-Host "Output saved to: $outFile"
