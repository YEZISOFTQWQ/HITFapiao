from __future__ import annotations

from dataclasses import dataclass
from datetime import date
from decimal import Decimal, ROUND_HALF_UP
from pathlib import Path
import shutil

from .models import InvoiceItem


BASE_DATA_START_ROW = 5
BASE_DATA_END_ROW = 14
BASE_TOTAL_ROW = 15
ROWS_PER_PRINT_PAGE = 10
MONEY = Decimal("0.01")


class ExportError(RuntimeError):
    pass


@dataclass
class AcceptanceMetadata:
    unit_name: str = ""
    expense_card: str = ""


@dataclass
class ExportResult:
    workbook_path: Path
    worksheet_count: int
    item_count: int
    total: Decimal


def _excel_color(red: int, green: int, blue: int) -> int:
    return red + green * 256 + blue * 65536


RED_FILL = _excel_color(255, 199, 206)
YELLOW_FILL = _excel_color(255, 235, 156)


def _extend_detail_rows(sheet, item_count: int) -> tuple[int, int, int]:
    """在合计行前增加明细行，返回明细末行、合计行和备注末行。"""
    extra_rows = max(0, item_count - ROWS_PER_PRINT_PAGE)
    if extra_rows:
        insert_start = BASE_TOTAL_ROW
        insert_end = BASE_TOTAL_ROW + extra_rows - 1
        sheet.Rows(f"{insert_start}:{insert_end}").Insert(-4121)  # xlShiftDown

        source = sheet.Range(f"A{BASE_DATA_END_ROW}:F{BASE_DATA_END_ROW}")
        destination = sheet.Range(f"A{insert_start}:F{insert_end}")
        source.Copy()
        destination.PasteSpecial(-4122)  # xlPasteFormats
        sheet.Application.CutCopyMode = False
        source_height = sheet.Rows(BASE_DATA_END_ROW).RowHeight
        for row in range(insert_start, insert_end + 1):
            sheet.Rows(row).RowHeight = source_height

    data_end_row = max(BASE_DATA_END_ROW, BASE_DATA_START_ROW + item_count - 1)
    total_row = BASE_TOTAL_ROW + extra_rows
    notes_end_row = total_row + 3
    return data_end_row, total_row, notes_end_row


def _fill_sheet(
    sheet,
    items: list[InvoiceItem],
    metadata: AcceptanceMetadata,
    acceptance_date: date,
    high_value_suppliers: set[str],
):
    data_end_row, total_row, notes_end_row = _extend_detail_rows(sheet, len(items))

    sheet.Range("C2").Value = acceptance_date.year
    sheet.Range("D2").Value = f"年  {acceptance_date.month:02d} 月"
    sheet.Range("E2").Value = f"{acceptance_date.day:02d} 日"
    sheet.Range("D2:E2").ShrinkToFit = True
    sheet.Range("A3").Value = f"单位名称：{metadata.unit_name}" if metadata.unit_name else "单位名称："
    sheet.Range("E3").Value = f"经费卡：{metadata.expense_card}" if metadata.expense_card else "经费卡："

    data_range = sheet.Range(f"A{BASE_DATA_START_ROW}:F{data_end_row}")
    data_range.UnMerge()
    data_range.ClearContents()
    data_range.VerticalAlignment = -4108  # xlCenter

    for offset, item in enumerate(items):
        row = BASE_DATA_START_ROW + offset
        sheet.Cells(row, 1).Value = item.name
        sheet.Cells(row, 2).Value = item.unit
        sheet.Cells(row, 3).Value = float(item.quantity)
        unit_price = item.unit_price_including_tax
        sheet.Cells(row, 4).Value = float(unit_price)
        sheet.Cells(row, 5).Value = float(item.amount_including_tax)
        sheet.Cells(row, 6).Value = item.supplier

        if item.red_flags:
            sheet.Range(f"A{row}:E{row}").Interior.Color = RED_FILL
            if item.supplier not in high_value_suppliers:
                sheet.Cells(row, 6).Interior.Color = RED_FILL

        quantity_cell = sheet.Cells(row, 3)
        quantity_cell.NumberFormat = "0" if item.quantity == item.quantity.to_integral_value() else "0.####"
        price_cell = sheet.Cells(row, 4)
        if abs(unit_price) >= Decimal("1000"):
            price_cell.NumberFormat = "0.00"
        elif unit_price == unit_price.quantize(MONEY):
            price_cell.NumberFormat = "0.00"
        else:
            price_cell.NumberFormat = "0.######"

        if len(item.name) > 18:
            name_cell = sheet.Cells(row, 1)
            name_cell.WrapText = True
            name_cell.ShrinkToFit = False
            name_cell.Font.Size = 8
            sheet.Rows(row).RowHeight = 28

    sheet.Range(f"E{BASE_DATA_START_ROW}:E{total_row}").NumberFormat = "0.00"
    sheet.Range(f"A{BASE_DATA_START_ROW}:A{data_end_row}").ShrinkToFit = True
    sheet.Range(f"F{BASE_DATA_START_ROW}:F{data_end_row}").ShrinkToFit = True
    sheet.Range(f"A{BASE_DATA_START_ROW}:A{data_end_row}").HorizontalAlignment = -4131  # xlLeft
    sheet.Range(f"B{BASE_DATA_START_ROW}:C{data_end_row}").HorizontalAlignment = -4108  # xlCenter
    sheet.Range(f"D{BASE_DATA_START_ROW}:E{total_row}").HorizontalAlignment = -4152  # xlRight
    sheet.Range(f"F{BASE_DATA_START_ROW}:F{data_end_row}").HorizontalAlignment = -4131  # xlLeft

    # 同一供应商连续出现时合并供应商单元格，保持模板的清晰度。
    start = 0
    while start < len(items):
        end = start
        while (
            end + 1 < len(items)
            and items[end + 1].supplier == items[start].supplier
            and (end + 1) // ROWS_PER_PRINT_PAGE == start // ROWS_PER_PRINT_PAGE
        ):
            end += 1
        excel_start = BASE_DATA_START_ROW + start
        excel_end = BASE_DATA_START_ROW + end
        if excel_end > excel_start:
            area = sheet.Range(f"F{excel_start}:F{excel_end}")
            area.Merge()
            area.Value = items[start].supplier
            area.VerticalAlignment = -4108
            area.HorizontalAlignment = -4131
            area.ShrinkToFit = True
        else:
            area = sheet.Range(f"F{excel_start}:F{excel_end}")
        if items[start].supplier in high_value_suppliers:
            area.Interior.Color = YELLOW_FILL
        elif any(item.red_flags for item in items[start : end + 1]):
            area.Interior.Color = RED_FILL
        start = end + 1

    sheet.Range(f"E{total_row}").Formula = (
        f"=ROUND(SUM(E{BASE_DATA_START_ROW}:E{BASE_DATA_START_ROW + len(items) - 1}),2)"
    )

    sheet.PageSetup.PrintArea = f"$A$1:$F${notes_end_row}"
    sheet.PageSetup.PrintTitleRows = "$1:$4"
    sheet.PageSetup.Orientation = 1  # xlPortrait
    sheet.PageSetup.PaperSize = 9  # xlPaperA4
    sheet.PageSetup.Zoom = False
    sheet.PageSetup.FitToPagesWide = 1
    sheet.PageSetup.FitToPagesTall = False

    sheet.ResetAllPageBreaks()
    for offset in range(ROWS_PER_PRINT_PAGE, len(items), ROWS_PER_PRINT_PAGE):
        sheet.HPageBreaks.Add(sheet.Rows(BASE_DATA_START_ROW + offset))

    return total_row


def export_acceptance_forms(
    template_path: str | Path,
    output_path: str | Path,
    items: list[InvoiceItem],
    metadata: AcceptanceMetadata,
) -> ExportResult:
    if not items:
        raise ExportError("没有可写入验收单的材料明细")

    template = Path(template_path).resolve()
    output = Path(output_path).resolve()
    if not template.is_file():
        raise FileNotFoundError(f"模板不存在：{template}")
    if template.suffix.casefold() != ".xls":
        raise ExportError("当前版本要求使用 .xls 模板")
    if output.suffix.casefold() != ".xls":
        output = output.with_suffix(".xls")
    output.parent.mkdir(parents=True, exist_ok=True)
    shutil.copy2(template, output)

    try:
        import pythoncom
        import win32com.client
    except ImportError as exc:
        raise ExportError("缺少 pywin32，无法调用 Excel 写入旧版 .xls 模板") from exc

    excel = None
    workbook = None
    acceptance_date = date.today()
    supplier_totals: dict[str, Decimal] = {}
    for item in items:
        supplier_totals[item.supplier] = supplier_totals.get(item.supplier, Decimal("0")) + item.amount_including_tax
    high_value_suppliers = {
        supplier for supplier, total in supplier_totals.items() if total > Decimal("1000")
    }
    try:
        pythoncom.CoInitialize()
        excel = win32com.client.DispatchEx("Excel.Application")
        excel.Visible = False
        excel.DisplayAlerts = False
        excel.ScreenUpdating = False
        workbook = excel.Workbooks.Open(str(output))

        template_sheet = workbook.Worksheets(1)
        for index in range(workbook.Worksheets.Count, 1, -1):
            workbook.Worksheets(index).Delete()

        template_sheet.Name = "验收单"
        total_row = _fill_sheet(
            template_sheet,
            items,
            metadata,
            acceptance_date,
            high_value_suppliers,
        )

        excel.CalculateFull()
        workbook.Save()

        expected_total = sum((item.amount_including_tax for item in items), Decimal("0")).quantize(MONEY)
        actual_total = Decimal(str(template_sheet.Range(f"E{total_row}").Value or 0)).quantize(MONEY)
        if actual_total.quantize(MONEY) != expected_total:
            raise ExportError(f"导出后合计校验失败：表格 {actual_total}，明细 {expected_total}")

        return ExportResult(
            workbook_path=output,
            worksheet_count=1,
            item_count=len(items),
            total=expected_total,
        )
    except Exception as exc:
        if isinstance(exc, ExportError):
            raise
        raise ExportError(f"生成验收单失败：{exc}") from exc
    finally:
        if workbook is not None:
            workbook.Close(SaveChanges=False)
        if excel is not None:
            excel.Quit()
        try:
            pythoncom.CoUninitialize()
        except Exception:
            pass
