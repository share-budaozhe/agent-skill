#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
PDF → Markdown 快速提取脚本（pdf-to-markdown skill 的核心工具）。

自动检测每本 PDF 的文本层：带文本层的数字版提取为 markdown；
纯扫描版（绝大多数页面无文本）跳过并列入报告。

用法:
    python extract_text.py [--input DIR] [--output DIR] [--empty-threshold 0.8]

依赖: pypdfium2  (pip install pypdfium2)
"""
import argparse
import os
import re
import sys

import pypdfium2 as pdfium

DEFAULT_THRESHOLD = 0.8  # 页面无文本比例超过该值判定为扫描版


def classify_pdf(path, threshold):
    """返回 (总页数, 无文本页数, 是否扫描版)。"""
    pdf = pdfium.PdfDocument(path)
    n = len(pdf)
    empty = 0
    for i in range(n):
        text = pdf[i].get_textpage().get_text_bounded()
        if not text.strip():
            empty += 1
    pdf.close()
    is_scanned = (empty / n) > threshold if n else True
    return n, empty, is_scanned


def clean_text(raw):
    """规范空白：统一行尾、合并多余空行、保留段落结构。"""
    raw = raw.replace("\r\n", "\n").replace("\r", "\n")
    out = []
    for ln in raw.split("\n"):
        s = ln.strip()
        if not s:
            if out and out[-1] != "":
                out.append("")
            continue
        out.append(s)
    text = "\n".join(out)
    text = re.sub(r"\n{3,}", "\n\n", text)
    text = re.sub(r"[ \t]+", " ", text)
    return text.strip()


def convert_to_markdown(path, title):
    """转换单本 PDF 为 markdown 文本。"""
    pdf = pdfium.PdfDocument(path)
    n = len(pdf)
    parts = []
    for i in range(n):
        raw = pdf[i].get_textpage().get_text_bounded()
        cleaned = clean_text(raw)
        if cleaned:
            parts.append(f"<!-- page {i+1} -->\n\n{cleaned}")
        else:
            parts.append(f"<!-- page {i+1}（无文本/图像页） -->")
    pdf.close()
    return f"# {title}\n\n" + "\n\n---\n\n".join(parts)


def main():
    parser = argparse.ArgumentParser(description="PDF → Markdown 快速提取")
    parser.add_argument("--input", default=os.getcwd(), help="包含 PDF 的目录（默认当前目录）")
    parser.add_argument("--output", default=None, help="输出目录（默认 <input>/md）")
    parser.add_argument("--empty-threshold", type=float, default=DEFAULT_THRESHOLD,
                        help="页面无文本比例超过该值判定为扫描版（默认 0.8）")
    args = parser.parse_args()

    in_dir = os.path.abspath(args.input)
    out_dir = os.path.abspath(args.output) if args.output else os.path.join(in_dir, "md")
    os.makedirs(out_dir, exist_ok=True)

    pdfs = sorted(f for f in os.listdir(in_dir) if f.lower().endswith(".pdf"))
    if not pdfs:
        print("未找到 PDF 文件。")
        return

    converted = 0
    skipped = 0
    print(f"共发现 {len(pdfs)} 个 PDF，输出到 {out_dir}\n")
    for f in pdfs:
        path = os.path.join(in_dir, f)
        n, empty, is_scanned = classify_pdf(path, args.empty_threshold)
        if is_scanned:
            print(f"[扫描版/跳过] {f}  （{n} 页，{empty} 页无文本，需 OCR）")
            skipped += 1
            continue
        title = f[:-4] if f.lower().endswith(".pdf") else f
        out_path = os.path.join(out_dir, title + ".md")
        print(f"[转换] {f}  ({n} 页, {empty} 页无文本)...")
        text = convert_to_markdown(path, title)
        with open(out_path, "w", encoding="utf-8") as fh:
            fh.write(text)
        print(f"  → {os.path.basename(out_path)}  ({len(text)//1024} KB)\n")
        converted += 1

    print("===== 结果报告 =====")
    print(f"完成 {converted} 本（数字版），跳过 {skipped} 本（扫描版，需 OCR）")
    print("输出目录:", out_dir)


if __name__ == "__main__":
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")
    main()
