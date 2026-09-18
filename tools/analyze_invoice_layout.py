from __future__ import annotations

import argparse
import json
from pathlib import Path

import pdfplumber


def page_lines(page):
    words = page.extract_words(x_tolerance=2, y_tolerance=2, keep_blank_chars=False)
    lines = []
    for word in sorted(words, key=lambda item: (round(item["top"], 1), item["x0"])):
        if not lines or abs(lines[-1][0] - word["top"]) > 2.0:
            lines.append([word["top"], []])
        lines[-1][1].append(word)
    return [
        {
            "top": round(top, 1),
            "text": " | ".join(
                f'{word["text"]} @{word["x0"]:.1f}-{word["x1"]:.1f}' for word in items
            ),
        }
        for top, items in lines
    ]


def main() -> int:
    parser = argparse.ArgumentParser(description="导出发票 PDF 的文字和表格布局")
    parser.add_argument("invoice_dir", help="待分析的 PDF 文件夹")
    parser.add_argument("--output", default=".work/invoice_layout.json")
    args = parser.parse_args()

    root = Path(args.invoice_dir).resolve()
    result = {}
    for path in sorted(root.glob("*.pdf")):
        with pdfplumber.open(path) as pdf:
            result[path.name] = [
                {
                    "width": page.width,
                    "height": page.height,
                    "lines": page_lines(page),
                    "tables": page.extract_tables(),
                }
                for page in pdf.pages
            ]

    output = Path(args.output).resolve()
    output.parent.mkdir(parents=True, exist_ok=True)
    output.write_text(json.dumps(result, ensure_ascii=False, indent=2), encoding="utf-8")
    print(output)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
