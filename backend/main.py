from datetime import date
from pathlib import Path
from typing import Literal

from fastapi import FastAPI, HTTPException
from fastapi.middleware.cors import CORSMiddleware
from pydantic import BaseModel, ConfigDict, Field
from fastapi.staticfiles import StaticFiles


ReviewStatus = Literal["Pending", "Approved", "Flagged", "Rejected"]
RiskStatus = Literal["High", "Medium", "Low"]


class LineItem(BaseModel):
    description: str
    hours: float = Field(ge=0)
    hourly_rate: float = Field(ge=0)
    contract_cap: float = 150.0

    @property
    def amount(self) -> float:
        return self.hours * self.hourly_rate


class AnomalyReport(BaseModel):
    rule: str
    severity: RiskStatus
    explanation: str


class Invoice(BaseModel):
    id: str
    vendor: str
    invoice_number: str
    amount: float
    invoice_date: date
    line_items: list[LineItem]
    review_status: ReviewStatus = "Pending"
    risk_status: RiskStatus = "Low"
    anomalies: list[AnomalyReport] = Field(default_factory=list)


class ReviewRequest(BaseModel):
    model_config = ConfigDict(extra="forbid", strict=True)

    status: Literal["Approved", "Flagged", "Rejected"]


app = FastAPI(title="AuditTrail AI API", version="1.0.0")
app.add_middleware(
    CORSMiddleware,
    allow_origins=["http://localhost:5173"],
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)


@app.get("/healthz", include_in_schema=False)
def health_check() -> dict[str, str]:
    return {"status": "ok"}


_invoices = [
    Invoice(
        id="inv-2401",
        vendor="Northstar Analytics",
        invoice_number="NSA-2026-0418",
        amount=8730.00,
        invoice_date=date(2026, 9, 22),
        line_items=[
            LineItem(description="Quarterly risk model review", hours=42, hourly_rate=150),
            LineItem(description="Executive findings workshop", hours=18, hourly_rate=135),
        ],
    ),
    Invoice(
        id="inv-2402",
        vendor="Meridian Strategy Group",
        invoice_number="MSG-2026-0882",
        amount=23075.00,
        invoice_date=date(2026, 9, 24),
        line_items=[
            LineItem(description="Regulatory controls assessment", hours=95, hourly_rate=185),
            LineItem(description="Senior partner review", hours=25, hourly_rate=220),
        ],
    ),
    Invoice(
        id="inv-2403",
        vendor="Clearwater Data Systems",
        invoice_number="CDS-2026-0197",
        amount=10580.00,
        invoice_date=date(2026, 9, 25),
        line_items=[
            LineItem(description="Ledger data migration", hours=48, hourly_rate=160),
            LineItem(description="Reconciliation support", hours=20, hourly_rate=145),
        ],
    ),
    Invoice(
        id="inv-2404",
        vendor="Northstar Analytics",
        invoice_number="CDS-2026-0197",
        amount=5500.00,
        invoice_date=date(2026, 9, 26),
        line_items=[
            LineItem(description="Model monitoring addendum", hours=28, hourly_rate=145),
            LineItem(description="Documentation update", hours=12, hourly_rate=120),
        ],
    ),
]


def evaluate_invoice(invoice: Invoice) -> Invoice:
    reports: list[AnomalyReport] = []
    for item in invoice.line_items:
        if item.hourly_rate > item.contract_cap:
            reports.append(
                AnomalyReport(
                    rule="Hourly rate exceeds contract cap",
                    severity="High",
                    explanation=(
                        f"{item.description} is billed at ${item.hourly_rate:.2f}/hr, "
                        f"above the contracted ${item.contract_cap:.2f}/hr cap."
                    ),
                )
            )

    duplicates = [
        candidate
        for candidate in _invoices
        if candidate.id != invoice.id
        and candidate.invoice_number == invoice.invoice_number
    ]
    if duplicates:
        reports.append(
            AnomalyReport(
                rule="Duplicate invoice number",
                severity="Medium",
                explanation=(
                    f"Invoice number {invoice.invoice_number} also appears on "
                    f"{duplicates[0].vendor}'s invoice {duplicates[0].id}."
                ),
            )
        )

    invoice.anomalies = reports
    invoice.risk_status = (
        "High"
        if any(report.severity == "High" for report in reports)
        else "Medium"
        if reports
        else "Low"
    )
    return invoice


@app.get("/api/invoices", response_model=list[Invoice])
def list_invoices() -> list[Invoice]:
    return [evaluate_invoice(invoice.model_copy(deep=True)) for invoice in _invoices]


@app.post("/api/invoices/{invoice_id}/review", response_model=Invoice)
def review_invoice(invoice_id: str, request: ReviewRequest) -> Invoice:
    invoice = next((item for item in _invoices if item.id == invoice_id), None)
    if invoice is None:
        raise HTTPException(status_code=404, detail="Invoice not found")
    invoice.review_status = request.status
    return evaluate_invoice(invoice.model_copy(deep=True))


static_dir = Path(__file__).resolve().parent / "static"
if static_dir.is_dir():
    app.mount("/", StaticFiles(directory=static_dir, html=True), name="frontend")