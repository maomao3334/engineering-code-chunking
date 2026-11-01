# -*- coding: utf-8 -*-
"""PDF 体检：有没有文字层？质量能不能直接用？页数/版面/表格情况。

决定走哪条路：
  有文字层且干净  → 直接抽取，跳过 OCR（省掉最贵的一步）
  无文字层/乱码    → 必须 OCR（双遍 300/600dpi 逐行择优）
"""
import sys
from pathlib import Path

import fitz

pdf = sys.argv[1]
doc = fitz.open(pdf)
print(f"页数: {doc.page_count}")
print(f"加密: {doc.is_encrypted}")
print(f"元数据: {doc.metadata.get('title')!r} / {doc.metadata.get('producer')!r}")

total_chars = 0
sample = []
for i in range(doc.page_count):
    t = doc[i].get_text()
    total_chars += len(t.strip())
    if i < 3 or (i in (5, 8, 10)):
        sample.append((i, t))

print(f"\n全文字符数: {total_chars}")
print(f"平均每页: {total_chars / max(doc.page_count,1):.0f} 字")

print("\n=== 抽样页内容 ===")
for i, t in sample:
    head = t.strip().replace("\n", " | ")[:220]
    print(f"\n--- p{i+1} ({len(t.strip())}字) ---")
    print(head if head else "【空——该页无文字层，是扫描图】")

# 判断是否文字层
avg = total_chars / max(doc.page_count, 1)
print("\n" + "=" * 60)
if avg > 300:
    print("判定：有文字层，可直接抽取（无需 OCR）")
elif avg > 50:
    print("判定：文字层稀疏（可能是双层PDF或含大量图），需逐页判断")
else:
    print("判定：基本无文字层 → 必须 OCR")

# 页面尺寸与旋转
p0 = doc[0]
print(f"页面尺寸: {p0.rect.width:.0f} x {p0.rect.height:.0f} pt, 旋转 {p0.rotation}")
doc.close()
