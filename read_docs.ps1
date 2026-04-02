Add-Type -AssemblyName Microsoft.Office.Interop.Word
$word = New-Object -ComObject Word.Application
$word.Visible = $false

$files = @(
    "D:\解压文件夹\2026招生宣传资料\学院简介+专业介绍\计算机科学与技术专业介绍2026.docx",
    "D:\解压文件夹\2026招生宣传资料\学院简介+专业介绍\软件工程专业介绍-2026.docx",
    "D:\解压文件夹\2026招生宣传资料\学院简介+专业介绍\电子信息工程专业-专业介绍2026.docx",
    "D:\解压文件夹\2026招生宣传资料\学院简介+专业介绍\人工智能专业介绍（2026）.docx",
    "D:\解压文件夹\2026招生宣传资料\学院简介+专业介绍\数据科学与大数据技术专业介绍2026.docx"
)

$output = @()
foreach ($file in $files) {
    $name = Split-Path $file -Leaf
    $output += "===== $name ====="
    $doc = $word.Documents.Open($file)
    $output += $doc.Content.Text
    $doc.Close($false)
    $output += "`n"
}

$word.Quit()
$output -join "`n"
