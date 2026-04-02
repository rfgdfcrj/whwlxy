@echo off
chcp 65001 >nul
setlocal enabledelayedexpansion

set "src_dir=D:\解压文件夹\2026招生宣传资料\学院简介+专业介绍"
set "out_dir=D:\解压文件夹\2026招生宣传资料\extracted"

if not exist "%out_dir%" mkdir "%out_dir%"

for %%f in ("%src_dir%\*.docx") do (
    echo Processing: %%~nxf
    set "filename=%%~nxf"
    set "name_only=%%~nf"
    
    powershell -Command "Add-Type -AssemblyName System.IO.Compression.FileSystem; [System.IO.Compression.ZipFile]::ExtractToDirectory('!src_dir!\!filename!', 'D:\解压文件夹\2026招生宣传资料\temp_extract')"
    
    powershell -Command "[xml]$xml = Get-Content 'D:\解压文件夹\2026招生宣传资料\temp_extract\word\document.xml' -Encoding UTF8; $ns = @{w = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'}; $texts = $xml.SelectNodes('//w:t', $ns) | ForEach-Object { $_.InnerText }; $texts -join '' | Out-File 'D:\解压文件夹\2026招生宣传资料\extracted\!name_only!.txt' -Encoding UTF8"
    
    rmdir /s /q "D:\解压文件夹\2026招生宣传资料\temp_extract"
)

echo Done!
pause
