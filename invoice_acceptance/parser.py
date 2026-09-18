from __future__ import annotations

from collections import Counter
from datetime import date
from decimal import Decimal, InvalidOperation, ROUND_HALF_UP
from pathlib import Path
import re

import pdfplumber

from .models import InvoiceItem, InvoiceRecord, ParseBatchResult, ParseFailure


MONEY = Decimal("0.01")
EXPECTED_BUYER_NAME = "哈尔滨工业大学"
EXPECTED_BUYER_TAX_ID = "12100000400000456B"


class InvoiceParseError(ValueError):
    pass


def _decimal(value: str, field_name: str) -> Decimal:
    cleaned = value.replace(",", "").replace("¥", "").replace("￥", "").strip()
    cleaned = cleaned.replace("—", "0").replace("-", "0")
    match = re.search(r"-?\d+(?:\.\d+)?", cleaned)
    if not match:
        raise InvoiceParseError(f"无法识别{field_name}：{value!r}")
    try:
        return Decimal(match.group())
    except InvalidOperation as exc:
        raise InvoiceParseError(f"无法识别{field_name}：{value!r}") from exc


def _natural_key(path: Path) -> list[object]:
    parts = re.split(r"(\d+)", path.name)
    return [int(part) if part.isdigit() else part.casefold() for part in parts]


def _group_lines(words: list[dict[str, object]], tolerance: float = 3.0):
    groups: list[tuple[float, list[dict[str, object]]]] = []
    for word in sorted(words, key=lambda item: (float(item["top"]), float(item["x0"]))):
        top = float(word["top"])
        if not groups or abs(groups[-1][0] - top) > tolerance:
            groups.append((top, [word]))
        else:
            groups[-1][1].append(word)
    return groups


def _line_text(words: list[dict[str, object]]) -> str:
    return "".join(str(word["text"]) for word in sorted(words, key=lambda item: float(item["x0"])))


def _join_name(tokens: list[dict[str, object]]) -> str:
    ordered = sorted(tokens, key=lambda item: float(item["x0"]))
    value = " ".join(str(token["text"]).strip() for token in ordered if str(token["text"]).strip())
    value = re.sub(r"(?<=[\u3400-\u9fff])\s+(?=[\u3400-\u9fff])", "", value)
    value = re.sub(r"\s+", " ", value).strip()
    return value


def _append_continuation(existing: str, continuation: str) -> str:
    if not existing:
        return continuation
    if not continuation:
        return existing
    if re.search(r"[\u3400-\u9fff]$", existing) and re.match(r"^[\u3400-\u9fff]", continuation):
        return existing + continuation
    return f"{existing} {continuation}"


def _clean_item_name(value: str) -> str:
    value = re.sub(r"^\s*\*[^*]+\*", "", value)
    value = re.sub(r"\s+", " ", value).strip()
    return value


def _keyword_from_filename(path: Path) -> str:
    value = re.sub(r"【.*?】", "", path.stem)
    value = re.sub(r"^\s*\d+\s*[.、_-]?\s*", "", value)
    return value.strip()


def _invoice_red_flags(is_standard_format: bool, buyer: str, buyer_tax_id: str) -> list[str]:
    flags: list[str] = []
    if not is_standard_format:
        flags.append("非标准发票格式")
    if buyer != EXPECTED_BUYER_NAME:
        flags.append(f"购买方名称错误：{buyer or '未识别'}")
    if buyer_tax_id != EXPECTED_BUYER_TAX_ID:
        flags.append(f"购买方纳税人识别号错误：{buyer_tax_id or '未识别'}")
    return flags


def _extract_parties(page) -> tuple[str, str, str, str]:
    candidates: list[tuple[str, str]] = []
    for table in page.extract_tables() or []:
        for row in table or []:
            for cell in row or []:
                if not cell:
                    continue
                compact = re.sub(r"\s+", "", cell)
                if "纳税人识别号" not in compact and "统一社会信用代码" not in compact:
                    continue
                name_match = re.search(
                    r"名称[:：](.+?)(?=统一社会信用代码|纳税人识别号|$)", compact
                )
                tax_match = re.search(
                    r"(?:统一社会信用代码/纳税人识别号|统一社会信用代码|纳税人识别号)[:：]([0-9A-Z]+)",
                    compact,
                    re.IGNORECASE,
                )
                if name_match and tax_match:
                    candidates.append((name_match.group(1).strip(), tax_match.group(1).upper()))

    if len(candidates) >= 2:
        return candidates[0][0], candidates[0][1], candidates[-1][0], candidates[-1][1]

    raise InvoiceParseError("未识别到购买方、销售方名称或纳税人识别号")


def _extract_header(text: str) -> tuple[str, date, Decimal]:
    number_match = re.search(r"发票号码\s*[:：]\s*(\d{8,20})", text)
    date_match = re.search(r"开票日期\s*[:：]\s*(\d{4})\s*年\s*(\d{1,2})\s*月\s*(\d{1,2})\s*日", text)
    total_match = re.search(r"[（(]小写[）)]\s*[¥￥]?\s*([\d,]+(?:\.\d+)?)", text)

    if not number_match:
        raise InvoiceParseError("未识别到发票号码")
    if not date_match:
        raise InvoiceParseError("未识别到开票日期")
    if not total_match:
        raise InvoiceParseError("未识别到价税合计（小写）")

    year, month, day = (int(value) for value in date_match.groups())
    return number_match.group(1), date(year, month, day), _decimal(total_match.group(1), "价税合计")


def _extract_items(
    page,
    source_file: Path,
    invoice_number: str,
    invoice_date: date,
    buyer: str,
    buyer_tax_id: str,
    supplier: str,
):
    words = page.extract_words(x_tolerance=2, y_tolerance=2, keep_blank_chars=False)
    groups = _group_lines(words)
    header_top = None
    total_top = None
    for top, line_words in groups:
        text = _line_text(line_words)
        if header_top is None and ("项目名称" in text or "货物或应税劳务" in text):
            header_top = top
            continue
        normalized = re.sub(r"\s+", "", text)
        if header_top is not None and top > header_top and ("合计" in normalized or normalized.startswith("合计")):
            total_top = top
            break

    if header_top is None or total_top is None:
        raise InvoiceParseError("未识别到发票商品表格")

    old_layout = "增值税电子普通发票" in (page.extract_text() or "")
    if old_layout:
        bounds = (235.0, 285.0, 340.0, 410.0, 480.0, 520.0)
    else:
        bounds = (180.0, 230.0, 305.0, 390.0, 445.0, 520.0)

    raw_items: list[dict[str, object]] = []
    for top, line_words in groups:
        if not (header_top + 5 < top < total_top - 2):
            continue

        columns: list[list[dict[str, object]]] = [[] for _ in range(7)]
        specification_tokens: list[dict[str, object]] = []
        for word in line_words:
            center = (float(word["x0"]) + float(word["x1"])) / 2
            if center < bounds[0]:
                if not old_layout and center >= 115.0:
                    specification_tokens.append(word)
                    continue
                index = 0
            elif center < bounds[1]:
                index = 1
            elif center < bounds[2]:
                index = 2
            elif center < bounds[3]:
                index = 3
            elif center < bounds[4]:
                index = 4
            elif center < bounds[5]:
                index = 5
            else:
                index = 6
            columns[index].append(word)

        name = _join_name(columns[0])
        specification = _join_name(specification_tokens)
        if name and specification:
            name = f"{name}（{specification}）"
        unit = _line_text(columns[1]).strip()
        quantity_text = _line_text(columns[2]).strip()
        price_text = _line_text(columns[3]).strip()
        amount_text = _line_text(columns[4]).strip()
        tax_rate = _line_text(columns[5]).strip()
        tax_text = _line_text(columns[6]).strip()

        if quantity_text and amount_text:
            try:
                raw_items.append(
                    {
                        "name": name,
                        "unit": unit,
                        "quantity": _decimal(quantity_text, "数量"),
                        "unit_price": _decimal(price_text, "单价"),
                        "amount": _decimal(amount_text, "金额"),
                        "tax_rate": tax_rate,
                        "tax": _decimal(tax_text, "税额"),
                    }
                )
            except InvoiceParseError:
                if name and raw_items:
                    raw_items[-1]["name"] = _append_continuation(str(raw_items[-1]["name"]), name)
                else:
                    raise
        elif name and raw_items:
            raw_items[-1]["name"] = _append_continuation(str(raw_items[-1]["name"]), name)

    if not raw_items:
        raise InvoiceParseError("商品表格中没有识别到有效明细")

    items: list[InvoiceItem] = []
    keyword = _keyword_from_filename(source_file)
    for raw in raw_items:
        amount_excluding_tax = Decimal(raw["amount"])
        tax_amount = Decimal(raw["tax"])
        amount_including_tax = (amount_excluding_tax + tax_amount).quantize(MONEY, rounding=ROUND_HALF_UP)
        original_name = _clean_item_name(str(raw["name"]))
        items.append(
            InvoiceItem(
                source_file=source_file,
                invoice_number=invoice_number,
                invoice_date=invoice_date,
                buyer=buyer,
                buyer_tax_id=buyer_tax_id,
                supplier=supplier,
                original_name=original_name,
                name=keyword or original_name,
                unit=str(raw["unit"]),
                quantity=Decimal(raw["quantity"]),
                unit_price_excluding_tax=Decimal(raw["unit_price"]),
                amount_excluding_tax=amount_excluding_tax,
                tax_rate=str(raw["tax_rate"]),
                tax_amount=tax_amount,
                amount_including_tax=amount_including_tax,
            )
        )
    return items


def parse_invoice(path: str | Path) -> InvoiceRecord:
    source_file = Path(path).resolve()
    with pdfplumber.open(source_file) as pdf:
        if not pdf.pages:
            raise InvoiceParseError("PDF 没有页面")
        page = pdf.pages[0]
        text = page.extract_text() or ""
        invoice_number, invoice_date, invoice_total = _extract_header(text)
        buyer, buyer_tax_id, supplier, _supplier_tax_id = _extract_parties(page)
        is_standard_format = "电子发票（普通发票）" in text
        items = _extract_items(
            page, source_file, invoice_number, invoice_date, buyer, buyer_tax_id, supplier
        )

    warnings: list[str] = []
    red_flags = _invoice_red_flags(is_standard_format, buyer, buyer_tax_id)
    warnings.extend(red_flags)
    for item in items:
        item.red_flags.extend(red_flags)
    detail_total = sum((item.amount_including_tax for item in items), Decimal("0")).quantize(MONEY)
    if detail_total != invoice_total.quantize(MONEY):
        warnings.append(f"明细含税合计 {detail_total} 与发票价税合计 {invoice_total} 不一致")

    filename_total = re.search(r"【\s*([\d,.]+)\s*】", source_file.name)
    if filename_total:
        expected = _decimal(filename_total.group(1), "文件名金额").quantize(MONEY)
        if expected != invoice_total.quantize(MONEY):
            warnings.append(f"文件名金额 {expected} 与发票价税合计 {invoice_total} 不一致")

    return InvoiceRecord(
        source_file=source_file,
        invoice_number=invoice_number,
        invoice_date=invoice_date,
        buyer=buyer,
        buyer_tax_id=buyer_tax_id,
        supplier=supplier,
        is_standard_format=is_standard_format,
        invoice_total=invoice_total.quantize(MONEY),
        items=items,
        red_flags=red_flags,
        warnings=warnings,
    )


def parse_invoice_directory(directory: str | Path) -> ParseBatchResult:
    root = Path(directory).resolve()
    if not root.is_dir():
        raise FileNotFoundError(f"发票目录不存在：{root}")

    pdfs = sorted(
        (path for path in root.rglob("*") if path.is_file() and path.suffix.casefold() == ".pdf"),
        key=_natural_key,
    )
    if not pdfs:
        raise FileNotFoundError(f"发票目录中没有 PDF：{root}")

    invoices: list[InvoiceRecord] = []
    failures: list[ParseFailure] = []
    for path in pdfs:
        try:
            invoices.append(parse_invoice(path))
        except Exception as exc:
            failures.append(ParseFailure(source_file=path, error=str(exc)))

    warnings: list[str] = []
    counts = Counter(invoice.invoice_number for invoice in invoices)
    duplicate_numbers = sorted(number for number, count in counts.items() if count > 1)
    if duplicate_numbers:
        warnings.append("检测到重复发票号码：" + "、".join(duplicate_numbers))

    buyers = sorted({invoice.buyer for invoice in invoices})
    if len(buyers) > 1:
        warnings.append("发票购买方不一致：" + "、".join(buyers))

    supplier_totals: dict[str, Decimal] = {}
    for invoice in invoices:
        for item in invoice.items:
            supplier_totals[item.supplier] = supplier_totals.get(item.supplier, Decimal("0")) + item.amount_including_tax
    for supplier, total in sorted(supplier_totals.items()):
        if total > Decimal("1000"):
            warnings.append(f"供应商累计金额超过1000元：{supplier}，合计 {total.quantize(MONEY)} 元")

    return ParseBatchResult(invoices=invoices, failures=failures, warnings=warnings)
