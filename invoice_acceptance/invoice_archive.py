from __future__ import annotations

from contextlib import ExitStack
from dataclasses import dataclass, field
from decimal import Decimal, ROUND_HALF_UP
from pathlib import Path
import re
import shutil

from pypdf import PageObject, PdfReader, PdfWriter, Transformation

from .models import InvoiceItem


A4_WIDTH = 595.2756
A4_HEIGHT = 841.8898
PAGE_MARGIN = 24.0
SLOT_GAP = 18.0
MONEY = Decimal("0.01")
COMBINED_PDF_NAME = "发票打印汇总.pdf"

# 标准归档文件名：序号.材料名称[金额].pdf。序号按整数比较，允许缺号和同号。
ARCHIVED_INVOICE_PATTERN = re.compile(
    r"^(?P<sequence>\d+)\.(?P<name>.+)\[(?P<amount>\d+(?:\.\d{1,2})?)\]\.pdf$",
    re.IGNORECASE,
)


class InvoiceArchiveError(RuntimeError):
    pass


@dataclass(frozen=True)
class ArchivedInvoice:
    path: Path
    sequence: int
    material_name: str
    amount: Decimal


@dataclass
class InvoiceArchiveResult:
    folder_path: Path
    combined_pdf_path: Path
    invoice_count: int
    source_page_count: int
    output_page_count: int
    excluded_files: list[Path] = field(default_factory=list)


def _parse_archived_invoice(path: Path) -> ArchivedInvoice | None:
    match = ARCHIVED_INVOICE_PATTERN.fullmatch(path.name)
    if not match:
        return None
    sequence = int(match.group("sequence"))
    if sequence < 1:
        return None
    return ArchivedInvoice(
        path=path,
        sequence=sequence,
        material_name=match.group("name"),
        amount=Decimal(match.group("amount")),
    )


def discover_archived_invoices(
    folder_path: str | Path,
    combined_pdf_path: str | Path | None = None,
) -> tuple[list[ArchivedInvoice], list[Path]]:
    """读取标准命名发票并按真实序号大小排序；缺号和同号都不会丢失。"""
    folder = Path(folder_path).resolve()
    if not folder.is_dir():
        raise FileNotFoundError(f"发票归档文件夹不存在：{folder}")

    output = Path(combined_pdf_path).resolve() if combined_pdf_path else (folder / COMBINED_PDF_NAME)
    invoices: list[ArchivedInvoice] = []
    excluded: list[Path] = []
    for path in folder.iterdir():
        if not path.is_file() or path.suffix.casefold() != ".pdf":
            continue
        if path.resolve() == output:
            continue
        invoice = _parse_archived_invoice(path)
        if invoice is None:
            excluded.append(path)
        else:
            invoices.append(invoice)

    invoices.sort(key=lambda invoice: (invoice.sequence, invoice.path.name.casefold(), invoice.path.name))
    excluded.sort(key=lambda path: (path.name.casefold(), path.name))
    return invoices, excluded


def _fit_page_into_slot(source_page, target_page: PageObject, slot_index: int) -> None:
    if int(source_page.get("/Rotate", 0) or 0) % 360:
        source_page.transfer_rotation_to_content()

    box = source_page.cropbox
    left = float(box.left)
    bottom = float(box.bottom)
    width = float(box.width)
    height = float(box.height)
    if width <= 0 or height <= 0:
        raise InvoiceArchiveError("发票 PDF 包含无效页面尺寸")

    slot_height = (A4_HEIGHT - PAGE_MARGIN * 2 - SLOT_GAP) / 2
    slot_width = A4_WIDTH - PAGE_MARGIN * 2
    scale = min(slot_width / width, slot_height / height)
    rendered_width = width * scale
    rendered_height = height * scale
    target_x = PAGE_MARGIN + (slot_width - rendered_width) / 2
    slot_bottom = PAGE_MARGIN if slot_index == 1 else PAGE_MARGIN + slot_height + SLOT_GAP
    target_y = slot_bottom + (slot_height - rendered_height) / 2

    transform = (
        Transformation()
        .translate(tx=-left, ty=-bottom)
        .scale(sx=scale, sy=scale)
        .translate(tx=target_x, ty=target_y)
    )
    target_page.merge_transformed_page(source_page, transform, expand=False)


def combine_archived_invoices(
    folder_path: str | Path,
    output_path: str | Path | None = None,
) -> InvoiceArchiveResult:
    """将归档文件夹中的标准命名发票按序号合成 A4 竖版双联打印 PDF。"""
    folder = Path(folder_path).resolve()
    output = Path(output_path).resolve() if output_path else (folder / COMBINED_PDF_NAME).resolve()
    invoices, excluded = discover_archived_invoices(folder, output)
    if not invoices:
        raise InvoiceArchiveError(
            "没有符合“序号.材料名称[金额].pdf”格式的发票；非标准命名文件已排除"
        )

    output.parent.mkdir(parents=True, exist_ok=True)
    writer = PdfWriter()
    target_page: PageObject | None = None
    source_page_count = 0

    try:
        with ExitStack() as stack:
            for invoice in invoices:
                stream = stack.enter_context(invoice.path.open("rb"))
                reader = PdfReader(stream)
                if reader.is_encrypted:
                    try:
                        reader.decrypt("")
                    except Exception as exc:
                        raise InvoiceArchiveError(f"无法读取加密发票：{invoice.path.name}") from exc
                if not reader.pages:
                    raise InvoiceArchiveError(f"发票 PDF 没有页面：{invoice.path.name}")
                for source_page in reader.pages:
                    slot_index = source_page_count % 2
                    if slot_index == 0:
                        target_page = PageObject.create_blank_page(width=A4_WIDTH, height=A4_HEIGHT)
                    assert target_page is not None
                    _fit_page_into_slot(source_page, target_page, slot_index)
                    source_page_count += 1
                    if slot_index == 1:
                        writer.add_page(target_page)
                        target_page = None

            if target_page is not None:
                writer.add_page(target_page)

            writer.add_metadata(
                {
                    "/Title": "发票打印汇总",
                    "/Subject": "A4 竖版，每页两张发票，按文件名前缀数字排序",
                    "/Creator": "发票批量生成材料验收单",
                }
            )
            temporary = output.with_name(f".{output.name}.tmp")
            writer.write(temporary)
            temporary.replace(output)
    except InvoiceArchiveError:
        raise
    except Exception as exc:
        raise InvoiceArchiveError(f"生成发票打印汇总失败：{exc}") from exc

    return InvoiceArchiveResult(
        folder_path=folder,
        combined_pdf_path=output,
        invoice_count=len(invoices),
        source_page_count=source_page_count,
        output_page_count=len(writer.pages),
        excluded_files=excluded,
    )


def _safe_material_name(value: str) -> str:
    cleaned = re.sub(r'[<>:"/\\|?*\x00-\x1f]', "_", value)
    cleaned = re.sub(r"\s+", " ", cleaned).strip(" .")
    return (cleaned or "材料")[:100].rstrip(" .")


def _group_items_by_source(items: list[InvoiceItem]) -> list[tuple[Path, list[InvoiceItem]]]:
    grouped: dict[Path, list[InvoiceItem]] = {}
    for item in items:
        source = item.source_file.resolve()
        grouped.setdefault(source, []).append(item)
    return list(grouped.items())


def archive_and_combine_invoices(
    items: list[InvoiceItem],
    workbook_output_path: str | Path,
) -> InvoiceArchiveResult:
    """按已确认材料名称归档原发票，并立即生成双联打印汇总。"""
    groups = _group_items_by_source(items)
    if not groups:
        raise InvoiceArchiveError("没有可归档的发票")

    workbook_path = Path(workbook_output_path).resolve()
    folder = workbook_path.parent / f"{workbook_path.stem}_发票"
    folder.mkdir(parents=True, exist_ok=True)

    if any(source.parent == folder.resolve() for source, _ in groups):
        raise InvoiceArchiveError("原发票目录不能与输出归档文件夹相同")

    # 刷新程序生成的标准命名发票，保留非标准命名文件供用户自行处理。
    old_invoices, _ = discover_archived_invoices(folder)
    for old_invoice in old_invoices:
        old_invoice.path.unlink()
    old_combined = folder / COMBINED_PDF_NAME
    if old_combined.exists():
        old_combined.unlink()

    for sequence, (source, source_items) in enumerate(groups, start=1):
        if not source.is_file():
            raise FileNotFoundError(f"原发票不存在：{source}")
        names: list[str] = []
        for item in source_items:
            name = item.name.strip()
            if name and name not in names:
                names.append(name)
        material_name = _safe_material_name("、".join(names))
        total = sum((item.amount_including_tax for item in source_items), Decimal("0"))
        total = total.quantize(MONEY, rounding=ROUND_HALF_UP)
        destination = folder / f"{sequence}.{material_name}[{total:.2f}].pdf"
        shutil.copy2(source, destination)

    return combine_archived_invoices(folder)
