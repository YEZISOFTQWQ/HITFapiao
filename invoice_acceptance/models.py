from __future__ import annotations

from dataclasses import dataclass, field
from datetime import date
from decimal import Decimal
from pathlib import Path


@dataclass
class InvoiceItem:
    source_file: Path
    invoice_number: str
    invoice_date: date
    buyer: str
    buyer_tax_id: str
    supplier: str
    original_name: str
    name: str
    unit: str
    quantity: Decimal
    unit_price_excluding_tax: Decimal
    amount_excluding_tax: Decimal
    tax_rate: str
    tax_amount: Decimal
    amount_including_tax: Decimal
    red_flags: list[str] = field(default_factory=list)
    warnings: list[str] = field(default_factory=list)

    @property
    def unit_price_including_tax(self) -> Decimal:
        if self.quantity == 0:
            return Decimal("0")
        return self.amount_including_tax / self.quantity

    def as_dict(self) -> dict[str, object]:
        return {
            "source_file": str(self.source_file),
            "invoice_number": self.invoice_number,
            "invoice_date": self.invoice_date.isoformat(),
            "buyer": self.buyer,
            "buyer_tax_id": self.buyer_tax_id,
            "supplier": self.supplier,
            "original_name": self.original_name,
            "name": self.name,
            "unit": self.unit,
            "quantity": str(self.quantity),
            "unit_price_excluding_tax": str(self.unit_price_excluding_tax),
            "unit_price_including_tax": str(self.unit_price_including_tax),
            "amount_excluding_tax": str(self.amount_excluding_tax),
            "tax_rate": self.tax_rate,
            "tax_amount": str(self.tax_amount),
            "amount_including_tax": str(self.amount_including_tax),
            "red_flags": list(self.red_flags),
            "warnings": list(self.warnings),
        }


@dataclass
class InvoiceRecord:
    source_file: Path
    invoice_number: str
    invoice_date: date
    buyer: str
    buyer_tax_id: str
    supplier: str
    is_standard_format: bool
    invoice_total: Decimal
    items: list[InvoiceItem]
    red_flags: list[str] = field(default_factory=list)
    warnings: list[str] = field(default_factory=list)

    def as_dict(self) -> dict[str, object]:
        return {
            "source_file": str(self.source_file),
            "invoice_number": self.invoice_number,
            "invoice_date": self.invoice_date.isoformat(),
            "buyer": self.buyer,
            "buyer_tax_id": self.buyer_tax_id,
            "supplier": self.supplier,
            "is_standard_format": self.is_standard_format,
            "invoice_total": str(self.invoice_total),
            "red_flags": list(self.red_flags),
            "warnings": list(self.warnings),
            "items": [item.as_dict() for item in self.items],
        }


@dataclass
class ParseFailure:
    source_file: Path
    error: str

    def as_dict(self) -> dict[str, str]:
        return {"source_file": str(self.source_file), "error": self.error}


@dataclass
class ParseBatchResult:
    invoices: list[InvoiceRecord]
    failures: list[ParseFailure]
    warnings: list[str] = field(default_factory=list)

    @property
    def items(self) -> list[InvoiceItem]:
        return [item for invoice in self.invoices for item in invoice.items]

    @property
    def total(self) -> Decimal:
        return sum((item.amount_including_tax for item in self.items), Decimal("0"))

    @property
    def supplier_totals(self) -> dict[str, Decimal]:
        totals: dict[str, Decimal] = {}
        for item in self.items:
            totals[item.supplier] = totals.get(item.supplier, Decimal("0")) + item.amount_including_tax
        return totals

    @property
    def high_value_suppliers(self) -> set[str]:
        return {supplier for supplier, total in self.supplier_totals.items() if total > Decimal("1000")}

    def as_dict(self) -> dict[str, object]:
        return {
            "invoice_count": len(self.invoices),
            "item_count": len(self.items),
            "total": str(self.total),
            "supplier_totals": {supplier: str(total) for supplier, total in self.supplier_totals.items()},
            "high_value_suppliers": sorted(self.high_value_suppliers),
            "warnings": list(self.warnings),
            "failures": [failure.as_dict() for failure in self.failures],
            "invoices": [invoice.as_dict() for invoice in self.invoices],
        }
