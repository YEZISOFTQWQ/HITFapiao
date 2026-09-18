from __future__ import annotations

import argparse
import json
import sys

from .exporter import AcceptanceMetadata, export_acceptance_forms
from .invoice_archive import archive_and_combine_invoices, combine_archived_invoices
from .parser import parse_invoice_directory


def _arguments() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="根据 PDF 发票批量生成哈尔滨工业大学材料验收单")
    parser.add_argument("--invoice-dir", help="PDF 发票目录")
    parser.add_argument("--template", help="材料验收单 .xls 模板")
    parser.add_argument("--output", help="输出 .xls 文件")
    parser.add_argument("--combine-existing", help="只读取已有的标准命名发票归档并生成打印汇总")
    parser.add_argument("--combined-output", help="已有发票打印汇总的输出 PDF；默认位于归档文件夹内")
    parser.add_argument("--unit-name", default="", help="单位名称；默认取发票购买方")
    parser.add_argument("--expense-card", default="", help="经费卡号")
    parser.add_argument("--allow-errors", action="store_true", help="有部分 PDF 解析失败时仍使用成功明细生成")
    parser.add_argument("--parse-only", action="store_true", help="只输出解析结果 JSON，不生成验收单")
    return parser.parse_args()


def main() -> int:
    args = _arguments()
    if args.combine_existing:
        try:
            archive_result = combine_archived_invoices(args.combine_existing, args.combined_output)
        except Exception as exc:
            print(f"生成打印汇总失败：{exc}", file=sys.stderr)
            return 2
        print(f"已生成发票打印汇总：{archive_result.combined_pdf_path}")
        print(
            f"标准命名发票 {archive_result.invoice_count} 张，原始页面 {archive_result.source_page_count} 页，"
            f"A4 输出 {archive_result.output_page_count} 页"
        )
        for path in archive_result.excluded_files:
            print(f"已排除非标准命名 PDF：{path.name}", file=sys.stderr)
        return 0

    if not args.invoice_dir:
        print("缺少必需参数：--invoice-dir", file=sys.stderr)
        return 2

    batch = parse_invoice_directory(args.invoice_dir)
    if args.parse_only:
        print(json.dumps(batch.as_dict(), ensure_ascii=False, indent=2))
        return 1 if batch.failures else 0

    missing = [
        option
        for option, value in (("--template", args.template), ("--output", args.output))
        if not value
    ]
    if missing:
        print(f"缺少必需参数：{', '.join(missing)}", file=sys.stderr)
        return 2

    if batch.failures and not args.allow_errors:
        for failure in batch.failures:
            print(f"解析失败：{failure.source_file.name}：{failure.error}", file=sys.stderr)
        print("存在解析失败的发票，未生成验收单。可修正文件或使用 --allow-errors。", file=sys.stderr)
        return 2
    if not batch.items:
        print("没有成功解析出材料明细。", file=sys.stderr)
        return 2

    buyers = sorted({invoice.buyer for invoice in batch.invoices})
    unit_name = args.unit_name or (buyers[0] if len(buyers) == 1 else "")
    metadata = AcceptanceMetadata(
        unit_name=unit_name,
        expense_card=args.expense_card,
    )
    result = export_acceptance_forms(
        template_path=args.template,
        output_path=args.output,
        items=batch.items,
        metadata=metadata,
    )
    archive_result = archive_and_combine_invoices(batch.items, result.workbook_path)

    for warning in batch.warnings:
        print(f"警告：{warning}", file=sys.stderr)
    for invoice in batch.invoices:
        for warning in invoice.warnings:
            print(f"警告：{invoice.source_file.name}：{warning}", file=sys.stderr)

    print(f"已生成：{result.workbook_path}")
    print(f"发票归档：{archive_result.folder_path}")
    print(f"发票打印汇总：{archive_result.combined_pdf_path}")
    print(
        f"发票 {len(batch.invoices)} 张，材料 {result.item_count} 条，"
        f"共 {result.worksheet_count} 个工作表，合计 {result.total} 元；"
        f"打印汇总 {archive_result.output_page_count} 页"
    )
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
