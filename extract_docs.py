# -*- coding: utf-8 -*-
import os
from docx import Document

base_dir = r"D:\解压文件夹\2026招生宣传资料"
docs_dir = os.path.join(base_dir, "学院简介+专业介绍")
output_file = os.path.join(base_dir, "专业介绍汇总.txt")

files = [
    "计算机科学与技术专业介绍2026.docx",
    "软件工程专业介绍-2026.docx",
    "电子信息工程专业-专业介绍2026.docx",
    "人工智能专业介绍（2026）.docx",
    "数据科学与大数据技术专业介绍2026.docx"
]

with open(output_file, 'w', encoding='utf-8') as out:
    for filename in files:
        filepath = os.path.join(docs_dir, filename)
        print(f"Processing: {filename}")
        
        try:
            doc = Document(filepath)
            text = []
            for para in doc.paragraphs:
                if para.text.strip():
                    text.append(para.text)
            
            out.write(f"===== {filename} =====\n")
            out.write("\n".join(text))
            out.write("\n\n")
        except Exception as e:
            print(f"Error processing {filename}: {e}")

print(f"Done! Output: {output_file}")
