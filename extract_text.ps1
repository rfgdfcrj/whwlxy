$inputDir = "D:\解压文件夹\2026招生宣传资料\学院简介+专业介绍"
$outputDir = "D:\解压文件夹\2026招生宣传资料\temp_docx"
$outFile = "D:\解压文件夹\2026招生宣传资料\专业介绍内容.txt"

$files = @(
    "计算机科学与技术专业介绍2026.docx",
    "软件工程专业介绍-2026.docx",
    "电子信息工程专业-专业介绍2026.docx",
    "人工智能专业介绍（2026）.docx",
    "数据科学与大数据技术专业介绍2026.docx"
)

$output = @()

foreach ($f in $files) {
    $srcPath = Join-Path $inputDir $f
    $zipPath = Join-Path $env:TEMP "temp_$([guid]::NewGuid().ToString()).zip"
    
    Copy-Item $srcPath $zipPath -Force
    
    $extractDir = Join-Path $outputDir ([System.IO.Path]::GetFileNameWithoutExtension($f))
    if (Test-Path $extractDir) { Remove-Item $extractDir -Recurse -Force }
    Expand-Archive -Path $zipPath -DestinationPath $extractDir -Force
    
    $xmlPath = Join-Path $extractDir "word\document.xml"
    [xml]$xml = Get-Content $xmlPath -Encoding UTF8
    
    $ns = @{w = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"}
    $texts = $xml.SelectNodes("//w:t", $ns) | ForEach-Object { $_.InnerText }
    $text = $texts -join ""
    
    Remove-Item $zipPath -Force
    
    $output += "===== $f ====="
    $output += $text
    $output += "`n"
}

$output -join "`n" | Out-File -FilePath $outFile -Encoding UTF8
Write-Output "Done. Output: $outFile"
