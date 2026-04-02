Add-Type -AssemblyName System.IO.Compression.FileSystem

$basePath = "D:\解压文件夹\2026招生宣传资料"
$docsPath = Join-Path $basePath "学院简介+专业介绍"
$outputPath = Join-Path $basePath "专业介绍汇总.txt"

$outLines = @()

$fileList = @(
    "计算机科学与技术专业介绍2026.docx"
    "软件工程专业介绍-2026.docx"
    "电子信息工程专业-专业介绍2026.docx"
    "人工智能专业介绍（2026）.docx"
    "数据科学与大数据技术专业介绍2026.docx"
)

foreach ($fileName in $fileList) {
    $filePath = Join-Path $docsPath $fileName
    Write-Host "Processing $fileName..."
    
    $extractPath = Join-Path $basePath "temp_extract"
    if (Test-Path $extractPath) {
        Remove-Item $extractPath -Recurse -Force
    }
    
    [System.IO.Compression.ZipFile]::ExtractToDirectory($filePath, $extractPath)
    
    $xmlPath = Join-Path $extractPath "word\document.xml"
    $xmlContent = Get-Content -Path $xmlPath -Encoding UTF8 -Raw
    
    $xml = [xml]$xmlContent
    $ns = New-Object System.Xml.XmlNamespaceManager($xml.NameTable)
    $ns.AddNamespace("w", "http://schemas.openxmlformats.org/wordprocessingml/2006/main")
    
    $textNodes = $xml.SelectNodes("//w:t", $ns)
    $text = ($textNodes | ForEach-Object { $_.InnerText }) -join ""
    
    $outLines += "===== $fileName ====="
    $outLines += $text
    $outLines += ""
    $outLines += ""
    
    Remove-Item $extractPath -Recurse -Force
}

$outLines -join "`n" | Out-File -FilePath $outputPath -Encoding UTF8
Write-Host "Done! Output: $outputPath"
