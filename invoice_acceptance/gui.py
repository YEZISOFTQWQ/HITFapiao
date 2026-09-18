from __future__ import annotations

from datetime import date
from decimal import Decimal, InvalidOperation
import os
from pathlib import Path
import sys
import tkinter as tk
from tkinter import filedialog, messagebox, ttk

from .exporter import AcceptanceMetadata, export_acceptance_forms
from .invoice_archive import archive_and_combine_invoices, combine_archived_invoices
from .models import InvoiceItem, ParseBatchResult
from .parser import parse_invoice_directory


def application_directory() -> Path:
    """源码运行时返回项目目录，打包后返回 EXE 所在目录。"""
    if getattr(sys, "frozen", False):
        return Path(sys.executable).resolve().parent
    return Path(__file__).resolve().parents[1]


class MaterialNameReviewDialog(tk.Toplevel):
    """用键盘 Enter 逐条确认材料名称的快捷窗口。"""

    def __init__(self, parent, items, confirmed_indices, high_value_suppliers, on_change):
        super().__init__(parent)
        self.items: list[InvoiceItem] = items
        self.confirmed_indices: set[int] = confirmed_indices
        self.high_value_suppliers: set[str] = high_value_suppliers
        self.on_change = on_change
        self.index = next((i for i in range(len(items)) if i not in confirmed_indices), 0)

        self.title("逐条确认材料名称")
        self.geometry("760x430")
        self.resizable(False, False)
        self.transient(parent)
        self.grab_set()
        self.protocol("WM_DELETE_WINDOW", self._close)

        self.progress = tk.StringVar()
        self.file_text = tk.StringVar()
        self.supplier_text = tk.StringVar()
        self.detail_text = tk.StringVar()
        self.original_text = tk.StringVar()
        self.warning_text = tk.StringVar()
        self.name_text = tk.StringVar()

        frame = ttk.Frame(self, padding=18)
        frame.pack(fill="both", expand=True)
        frame.columnconfigure(1, weight=1)

        ttk.Label(frame, textvariable=self.progress, font=("Microsoft YaHei UI", 12, "bold")).grid(
            row=0, column=0, columnspan=2, sticky="w", pady=(0, 12)
        )
        self._info_row(frame, 1, "发票文件", self.file_text)
        self._info_row(frame, 2, "供应商", self.supplier_text)
        self._info_row(frame, 3, "数量金额", self.detail_text)
        self._info_row(frame, 4, "发票原文", self.original_text)
        ttk.Label(frame, textvariable=self.warning_text, foreground="#B91C1C", wraplength=600).grid(
            row=5, column=1, sticky="w", pady=(3, 8)
        )

        ttk.Label(frame, text="验收单名称").grid(row=6, column=0, sticky="ne", padx=(0, 10), pady=8)
        self.name_entry = ttk.Entry(frame, textvariable=self.name_text, font=("Microsoft YaHei UI", 14))
        self.name_entry.grid(row=6, column=1, sticky="ew", pady=8)
        ttk.Label(
            frame,
            text="直接修改后按 Enter 确认下一条；Ctrl+O 打开当前发票。",
            foreground="#555555",
        ).grid(row=7, column=1, sticky="w", pady=(0, 14))

        buttons = ttk.Frame(frame)
        buttons.grid(row=8, column=0, columnspan=2, sticky="ew")
        ttk.Button(buttons, text="打开原发票（Ctrl+O）", command=self._open_invoice).pack(side="left")
        ttk.Button(buttons, text="使用发票原文", command=self._use_original).pack(side="left", padx=8)
        ttk.Button(buttons, text="上一条", command=self._previous).pack(side="right")
        self.confirm_button = ttk.Button(buttons, text="确认并下一条（Enter）", command=self._confirm_next)
        self.confirm_button.pack(side="right", padx=(8, 0))

        self.bind("<Return>", lambda _event: self._confirm_next())
        self.bind("<Control-o>", lambda _event: self._open_invoice())
        self.bind("<Control-O>", lambda _event: self._open_invoice())
        self._load_current()

    @staticmethod
    def _info_row(frame, row, title, variable):
        ttk.Label(frame, text=title).grid(row=row, column=0, sticky="ne", padx=(0, 10), pady=3)
        ttk.Label(frame, textvariable=variable, wraplength=600).grid(row=row, column=1, sticky="w", pady=3)

    def _current_warning(self, item: InvoiceItem) -> str:
        warnings = list(item.red_flags)
        if item.supplier in self.high_value_suppliers:
            warnings.append("该供应商累计发票金额超过 1000 元，将在验收单中标黄")
        return "；".join(warnings)

    def _load_current(self):
        item = self.items[self.index]
        self.progress.set(
            f"第 {self.index + 1}/{len(self.items)} 条    已确认 {len(self.confirmed_indices)}/{len(self.items)}"
        )
        self.file_text.set(item.source_file.name)
        self.supplier_text.set(item.supplier)
        self.detail_text.set(
            f"{item.quantity} {item.unit}，含税单价 {item.unit_price_including_tax:.4f} 元，金额 {item.amount_including_tax:.2f} 元"
        )
        self.original_text.set(item.original_name)
        self.warning_text.set(self._current_warning(item))
        self.name_text.set(item.name)
        remaining = len(self.items) - len(self.confirmed_indices)
        self.confirm_button.configure(text="确认并完成（Enter）" if remaining <= 1 else "确认并下一条（Enter）")
        self.name_entry.focus_set()
        self.name_entry.select_range(0, "end")

    def _use_original(self):
        self.name_text.set(self.items[self.index].original_name)
        self.name_entry.focus_set()
        self.name_entry.select_range(0, "end")

    def _open_invoice(self):
        try:
            os.startfile(str(self.items[self.index].source_file))
        except OSError as exc:
            messagebox.showerror("无法打开发票", str(exc), parent=self)

    def _previous(self):
        self.index = (self.index - 1) % len(self.items)
        self._load_current()

    def _confirm_next(self):
        value = self.name_text.get().strip()
        if not value:
            messagebox.showerror("名称不能为空", "请输入材料验收单中的材料名称。", parent=self)
            return
        self.items[self.index].name = value
        self.confirmed_indices.add(self.index)
        self.on_change()
        if len(self.confirmed_indices) == len(self.items):
            messagebox.showinfo("确认完成", "所有材料名称已确认。", parent=self)
            self._close()
            return
        for step in range(1, len(self.items) + 1):
            candidate = (self.index + step) % len(self.items)
            if candidate not in self.confirmed_indices:
                self.index = candidate
                break
        self._load_current()

    def _close(self):
        self.on_change()
        self.grab_release()
        self.destroy()


class InvoiceAcceptanceApp(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("发票批量生成材料验收单")
        self.geometry("1380x780")
        self.minsize(1120, 680)
        self.items: list[InvoiceItem] = []
        self.batch: ParseBatchResult | None = None
        self.confirmed_indices: set[int] = set()
        self.high_value_suppliers: set[str] = set()
        self._build_ui()

    def _build_ui(self):
        self.columnconfigure(0, weight=1)
        self.rowconfigure(2, weight=1)

        source = ttk.LabelFrame(self, text="文件")
        source.grid(row=0, column=0, sticky="ew", padx=12, pady=(12, 6))
        source.columnconfigure(1, weight=1)
        base_dir = application_directory()
        self.invoice_dir = tk.StringVar(value=str(base_dir.resolve()))
        self.template = tk.StringVar(value=str((base_dir / "模板.xls").resolve()))
        self.output = tk.StringVar(value=str((base_dir / "输出" / "材料验收单.xls").resolve()))
        self.archive_dir = tk.StringVar(value=str((base_dir / "输出" / "材料验收单_发票").resolve()))
        self._path_row(source, 0, "发票目录", self.invoice_dir, self._choose_invoice_dir)
        self._path_row(source, 1, "模板", self.template, self._choose_template)
        self._path_row(source, 2, "输出文件", self.output, self._choose_output)
        self._path_row(source, 3, "已有发票归档", self.archive_dir, self._choose_archive_dir)

        details = ttk.LabelFrame(self, text="验收单信息")
        details.grid(row=1, column=0, sticky="ew", padx=12, pady=6)
        for column in (1, 3, 5):
            details.columnconfigure(column, weight=1)
        self.fields = {"unit": tk.StringVar(), "card": tk.StringVar()}
        self.system_date = tk.StringVar(value=f"{date.today().isoformat()}（系统日期）")
        ttk.Label(details, text="验收日期").grid(row=0, column=0, sticky="e", padx=(8, 4), pady=6)
        ttk.Label(details, textvariable=self.system_date).grid(row=0, column=1, sticky="w", padx=(0, 18), pady=6)
        ttk.Label(details, text="单位名称").grid(row=0, column=2, sticky="e", padx=(8, 4), pady=6)
        ttk.Entry(details, textvariable=self.fields["unit"]).grid(row=0, column=3, sticky="ew", padx=(0, 18), pady=6)
        ttk.Label(details, text="经费卡").grid(row=0, column=4, sticky="e", padx=(8, 4), pady=6)
        ttk.Entry(details, textvariable=self.fields["card"]).grid(row=0, column=5, sticky="ew", padx=(0, 8), pady=6)

        table_frame = ttk.LabelFrame(self, text="解析明细（名称需逐条确认；表格仍支持双击修改）")
        table_frame.grid(row=2, column=0, sticky="nsew", padx=12, pady=6)
        table_frame.rowconfigure(0, weight=1)
        table_frame.columnconfigure(0, weight=1)
        columns = (
            "index", "original_name", "name", "unit", "quantity", "unit_price", "amount",
            "supplier", "invoice", "warning",
        )
        self.tree = ttk.Treeview(table_frame, columns=columns, show="headings", selectmode="browse")
        headings = {
            "index": "序号", "original_name": "发票原文", "name": "验收单名称", "unit": "单位",
            "quantity": "数量", "unit_price": "含税单价", "amount": "含税金额", "supplier": "供应商",
            "invoice": "发票号码", "warning": "校验提示",
        }
        widths = {
            "index": 45, "original_name": 210, "name": 130, "unit": 48, "quantity": 58,
            "unit_price": 82, "amount": 82, "supplier": 210, "invoice": 170, "warning": 260,
        }
        for column in columns:
            self.tree.heading(column, text=headings[column])
            anchor = "w" if column in {"original_name", "name", "supplier", "warning"} else "center"
            self.tree.column(column, width=widths[column], anchor=anchor)
        self.tree.tag_configure("red", background="#FFC7CE")
        self.tree.tag_configure("yellow", background="#FFEB9C")
        self.tree.grid(row=0, column=0, sticky="nsew")
        scrollbar = ttk.Scrollbar(table_frame, orient="vertical", command=self.tree.yview)
        scrollbar.grid(row=0, column=1, sticky="ns")
        self.tree.configure(yscrollcommand=scrollbar.set)
        self.tree.bind("<Double-1>", self._edit_cell)

        footer = ttk.Frame(self)
        footer.grid(row=3, column=0, sticky="ew", padx=12, pady=(6, 12))
        footer.columnconfigure(0, weight=1)
        self.status = tk.StringVar(value="请选择发票目录后解析。")
        ttk.Label(footer, textvariable=self.status).grid(row=0, column=0, sticky="w")
        ttk.Button(footer, text="解析发票", command=self.parse).grid(row=0, column=1, padx=6)
        ttk.Button(footer, text="逐条确认材料名称", command=self.review_names).grid(row=0, column=2, padx=6)
        ttk.Button(footer, text="已有发票生成打印PDF", command=self.combine_existing).grid(
            row=0, column=3, padx=6
        )
        ttk.Button(footer, text="生成验收单和发票归档", command=self.generate).grid(row=0, column=4)

    def _path_row(self, parent, row, label, variable, command):
        ttk.Label(parent, text=label).grid(row=row, column=0, sticky="e", padx=(8, 4), pady=5)
        ttk.Entry(parent, textvariable=variable).grid(row=row, column=1, sticky="ew", pady=5)
        ttk.Button(parent, text="选择…", command=command).grid(row=row, column=2, padx=8, pady=5)

    def _choose_invoice_dir(self):
        value = filedialog.askdirectory(initialdir=self.invoice_dir.get() or str(Path.cwd()))
        if value:
            self.invoice_dir.set(value)

    def _choose_template(self):
        value = filedialog.askopenfilename(filetypes=[("Excel 97-2003", "*.xls")])
        if value:
            self.template.set(value)

    def _choose_output(self):
        value = filedialog.asksaveasfilename(defaultextension=".xls", filetypes=[("Excel 97-2003", "*.xls")])
        if value:
            self.output.set(value)
            output = Path(value).resolve()
            self.archive_dir.set(str(output.parent / f"{output.stem}_发票"))

    def _choose_archive_dir(self):
        value = filedialog.askdirectory(initialdir=self.archive_dir.get() or str(Path.cwd()))
        if value:
            self.archive_dir.set(value)

    def parse(self):
        try:
            batch = parse_invoice_directory(self.invoice_dir.get())
        except Exception as exc:
            messagebox.showerror("解析失败", str(exc))
            return
        if batch.failures:
            details = "\n".join(f"{failure.source_file.name}：{failure.error}" for failure in batch.failures)
            messagebox.showerror("存在未解析发票", details)
            return

        self.batch = batch
        self.items = batch.items
        self.confirmed_indices.clear()
        self.high_value_suppliers = batch.high_value_suppliers
        self._refresh_tree()
        buyers = sorted({invoice.buyer for invoice in batch.invoices})
        if len(buyers) == 1 and not self.fields["unit"].get():
            self.fields["unit"].set(buyers[0])

        warning_count = len(batch.warnings) + sum(len(invoice.warnings) for invoice in batch.invoices)
        self._update_status(warning_count)
        messages = list(batch.warnings)
        for invoice in batch.invoices:
            messages.extend(f"{invoice.source_file.name}：{warning}" for warning in invoice.warnings)
        if messages:
            messagebox.showwarning("发票校验提示", "\n".join(messages))
        self.review_names()

    def _update_status(self, warning_count: int | None = None):
        if warning_count is None and self.batch:
            warning_count = len(self.batch.warnings) + sum(len(invoice.warnings) for invoice in self.batch.invoices)
        self.status.set(
            f"已解析 {len(self.batch.invoices) if self.batch else 0} 张发票、{len(self.items)} 条材料，"
            f"合计 {sum((item.amount_including_tax for item in self.items), Decimal('0')):.2f} 元；"
            f"名称已确认 {len(self.confirmed_indices)}/{len(self.items)}；提示 {warning_count or 0} 条。"
        )

    def _recalculate_supplier_alerts(self):
        totals: dict[str, Decimal] = {}
        for item in self.items:
            totals[item.supplier] = totals.get(item.supplier, Decimal("0")) + item.amount_including_tax
        self.high_value_suppliers = {supplier for supplier, total in totals.items() if total > Decimal("1000")}

    def _item_warning(self, item: InvoiceItem) -> str:
        warnings = list(item.red_flags)
        if item.supplier in self.high_value_suppliers:
            warnings.append("供应商累计金额超过1000元")
        return "；".join(warnings)

    def _refresh_tree(self):
        self.tree.delete(*self.tree.get_children())
        for index, item in enumerate(self.items, start=1):
            tags = ("red",) if item.red_flags else (("yellow",) if item.supplier in self.high_value_suppliers else ())
            self.tree.insert("", "end", iid=str(index - 1), tags=tags, values=(
                index, item.original_name, item.name, item.unit, str(item.quantity),
                f"{item.unit_price_including_tax:.4f}", f"{item.amount_including_tax:.2f}",
                item.supplier, item.invoice_number, self._item_warning(item),
            ))

    def review_names(self):
        if not self.items:
            messagebox.showinfo("尚未解析", "请先解析发票。")
            return
        MaterialNameReviewDialog(
            self, self.items, self.confirmed_indices, self.high_value_suppliers, self._review_changed
        )

    def _review_changed(self):
        self._refresh_tree()
        self._update_status()

    def _edit_cell(self, event):
        if self.tree.identify("region", event.x, event.y) != "cell":
            return
        item_id = self.tree.identify_row(event.y)
        column_id = self.tree.identify_column(event.x)
        column_index = int(column_id[1:]) - 1
        editable = {2: "name", 3: "unit", 4: "quantity", 5: "unit_price", 6: "amount", 7: "supplier"}
        if not item_id or column_index not in editable:
            return
        bbox = self.tree.bbox(item_id, column_id)
        current = self.tree.item(item_id, "values")[column_index]
        editor = ttk.Entry(self.tree)
        editor.insert(0, current)
        editor.select_range(0, "end")
        editor.place(x=bbox[0], y=bbox[1], width=bbox[2], height=bbox[3])
        editor.focus_set()

        def commit(_event=None):
            value = editor.get().strip()
            try:
                item_index = int(item_id)
                item = self.items[item_index]
                field = editable[column_index]
                if field == "quantity":
                    item.quantity = Decimal(value)
                elif field == "unit_price":
                    item.amount_including_tax = (Decimal(value) * item.quantity).quantize(Decimal("0.01"))
                elif field == "amount":
                    item.amount_including_tax = Decimal(value).quantize(Decimal("0.01"))
                else:
                    if not value:
                        raise ValueError("内容不能为空")
                    setattr(item, field, value)
                if field == "name":
                    self.confirmed_indices.add(item_index)
                self._recalculate_supplier_alerts()
                editor.destroy()
                self._refresh_tree()
                self._update_status()
            except (InvalidOperation, ValueError):
                messagebox.showerror("输入错误", "名称和供应商不能为空，数量、单价和金额必须是数字。")
                editor.focus_set()

        editor.bind("<Return>", commit)
        editor.bind("<FocusOut>", commit)
        editor.bind("<Escape>", lambda _event: editor.destroy())

    def generate(self):
        if not self.items:
            messagebox.showinfo("尚未解析", "请先解析发票。")
            return
        if len(self.confirmed_indices) != len(self.items):
            messagebox.showwarning("名称未确认", "请先逐条确认全部材料名称。")
            self.review_names()
            return
        try:
            self.system_date.set(f"{date.today().isoformat()}（系统日期）")
            metadata = AcceptanceMetadata(
                unit_name=self.fields["unit"].get().strip(),
                expense_card=self.fields["card"].get().strip(),
            )
            result = export_acceptance_forms(self.template.get(), self.output.get(), self.items, metadata)
            archive_result = archive_and_combine_invoices(self.items, result.workbook_path)
        except Exception as exc:
            messagebox.showerror("生成失败", str(exc))
            return
        self.archive_dir.set(str(archive_result.folder_path))
        self.status.set(
            f"已生成 {result.item_count} 条材料；归档 {archive_result.invoice_count} 张发票；"
            f"打印汇总 {archive_result.output_page_count} 页。"
        )
        messagebox.showinfo(
            "生成完成",
            f"验收单：{result.workbook_path}\n"
            f"发票归档：{archive_result.folder_path}\n"
            f"打印汇总：{archive_result.combined_pdf_path}",
        )

    def combine_existing(self):
        try:
            result = combine_archived_invoices(self.archive_dir.get())
        except Exception as exc:
            messagebox.showerror("生成打印PDF失败", str(exc))
            return

        self.status.set(
            f"按真实序号处理 {result.invoice_count} 张标准命名发票，"
            f"生成 {result.output_page_count} 页打印汇总；排除 {len(result.excluded_files)} 个文件。"
        )
        excluded_text = ""
        if result.excluded_files:
            excluded_text = "\n\n以下 PDF 因命名不符合“序号.材料名称[金额].pdf”而被排除：\n" + "\n".join(
                path.name for path in result.excluded_files
            )
        messagebox.showinfo(
            "打印PDF生成完成",
            f"读取发票：{result.invoice_count} 张\n"
            f"原始页数：{result.source_page_count} 页\n"
            f"输出页数：{result.output_page_count} 页\n"
            f"文件：{result.combined_pdf_path}"
            f"{excluded_text}",
        )


def main():
    app = InvoiceAcceptanceApp()
    app.mainloop()


if __name__ == "__main__":
    main()
