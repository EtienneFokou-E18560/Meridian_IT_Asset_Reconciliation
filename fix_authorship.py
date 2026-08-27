#!/usr/bin/env python3
"""Strip LLM-authorship tells from Meridian inputs and rebuild the packet zip."""
from __future__ import annotations

import csv
import zipfile
from datetime import date, timedelta
from pathlib import Path

from PIL import Image, ImageDraw, ImageFont
from reportlab.lib import colors
from reportlab.lib.pagesizes import letter
from reportlab.lib.styles import ParagraphStyle, getSampleStyleSheet
from reportlab.lib.units import inch
from reportlab.platypus import Image as RLImage
from reportlab.platypus import PageBreak, Paragraph, Preformatted, SimpleDocTemplate, Spacer

ROOT = Path(__file__).resolve().parent

NAMES = [
    "Priya Shah", "Marcus Ellison", "Elena Kovacs", "James Whitaker", "Aisha Rahman",
    "Colin Bergstrom", "Sofia Alvarez", "Derek Nguyen", "Hannah Cole", "Owen Patel",
    "Maya Fontaine", "Ryan Okonkwo", "Lauren Briggs", "Nathan Kim", "Camila Ortiz",
    "Patrick Gallagher", "Irene Cho", "Theo Barnett", "Jasmine Reed", "Luis Navarro",
    "Grace Lindholm", "Andre Baptiste", "Nina Kowalski", "Chris Daley", "Fatima Hussein",
    "Brendan Walsh", "Yuki Tanaka", "Megan Copeland", "Omar Haddad", "Stephanie Ruiz",
    "Keith Moreau", "Diana Voss", "Jordan Blake", "Amelia Grant", "Rafael Mendes",
    "Kate Sorenson", "Victor Lang", "Noelle Harper", "Samir Kapoor", "Holly Jensen",
    "Eric Vann", "Claudia Rossi", "Tyler Brooks", "Leah Okada", "Darnell Price",
    "Monica Ferreira", "Gavin Crowe", "Sana Iqbal", "Brett Holloway", "Paula Chen",
    "Isaac Moran", "Rebecca Sutter", "Kwame Asante", "Allison Pike", "Marco DeLuca",
    "Cynthia Hale", "Noah Frazier", "Rina Desai", "Todd Campbell", "Vanessa Ortega",
    "Liam Keane", "Bianca Silva", "Greg Hollis", "Naomi Berger", "Felix Moreno",
    "Sharon Quinlan", "Adrian Moss", "Kelsey Tran", "Malik Jordan", "Heidi Volkov",
    "Sean McCabe", "Anita Bose", "Drew Carmichael", "Lila Novak", "Curtis Bennett",
    "Pilar Santos", "Jonah Freedman", "Regina Lowe", "Wes Carvalho", "Tara Singh",
    "Elliott Park", "Carmen Diaz", "Hugh Talbot", "Mira Adelman", "Quincy Rhodes",
    "Sloane Richter", "Nolan Vega", "Ivy Chenoweth", "Percy Lambert", "Dana Ghosh",
    "Harlan Smith", "June Okoye", "Rory Phelps", "Cecilia Hart", "Emmett Boyle",
]


def font(size: int, bold: bool = False):
    names = (
        ["Arial Bold.ttf", "Arial Unicode.ttf"] if bold else ["Arial.ttf", "Arial Unicode.ttf"]
    )
    search = []
    if bold:
        search += [
            "/System/Library/Fonts/Supplemental/Arial Bold.ttf",
            "/Library/Fonts/Arial Bold.ttf",
        ]
    search += [
        "/System/Library/Fonts/Supplemental/Arial.ttf",
        "/Library/Fonts/Arial.ttf",
        "/System/Library/Fonts/Helvetica.ttc",
    ]
    for p in search:
        if Path(p).exists():
            try:
                return ImageFont.truetype(p, size)
            except OSError:
                continue
    return ImageFont.load_default()


def render_control_matrix() -> None:
    w, h = 1600, 900
    img = Image.new("RGB", (w, h), (248, 246, 241))  # warm paper, not cool AI-blue
    d = ImageDraw.Draw(img)
    # charcoal frame, not navy
    d.rectangle([18, 18, w - 19, h - 19], outline=(90, 74, 61), width=3)

    title_f = font(34, bold=True)
    sub_f = font(18)
    head_f = font(20, bold=True)
    cell_f = font(18)
    foot_f = font(16)

    d.text((48, 42), "Yanou & Partners IT Services — ITAM-001 Appendix A Control Matrix", font=title_f, fill=(52, 42, 36))
    d.text((48, 92), "Effective with Policy ITAM-001. Read with the narrative controls.", font=sub_f, fill=(92, 82, 74))

    headers = ["Control", "Required Evidence", "Failure Classification", "Owner"]
    rows = [
        ["Assignment", "Employee ID + tag + serial + acknowledgement", "Unverified custody", "IT Asset Specialist"],
        ["Transfer", "Transfer ID + approval + receipt", "Unapproved transfer", "IT Operations Manager"],
        ["Return", "Carrier acceptance + delivery + receiving scan", "Overdue / shipment mismatch", "IT Asset Specialist"],
        ["Disposal", "Retirement approval + sanitization + certificate", "Missing disposal evidence", "IT Asset Manager"],
        ["Financial", "PO + register + ledger agreement", "Inventory-to-ledger difference", "Finance Controller"],
    ]
    cols_x = [48, 280, 820, 1220, 1552]
    y0 = 150
    row_h = 92
    header_fill = (92, 74, 61)  # brown-charcoal, not navy/AI-blue
    alt = (241, 232, 220)

    d.rectangle([cols_x[0], y0, cols_x[-1], y0 + row_h], fill=header_fill)
    for i, hdr in enumerate(headers):
        d.text((cols_x[i] + 14, y0 + 30), hdr, font=head_f, fill=(255, 252, 246))

    for r, row in enumerate(rows):
        top = y0 + row_h * (r + 1)
        fill = alt if r % 2 else (252, 250, 246)
        d.rectangle([cols_x[0], top, cols_x[-1], top + row_h], fill=fill, outline=(210, 198, 184))
        for i, val in enumerate(row):
            d.text((cols_x[i] + 14, top + 32), val, font=cell_f, fill=(45, 38, 32))

    d.text(
        (48, h - 62),
        "Use with IT_asset_management_policy.pdf. Informal notes cannot override this evidence standard.",
        font=foot_f,
        fill=(110, 70, 48),
    )
    out = ROOT / "ITAM_control_matrix.png"
    img.save(out, "PNG")
    print("wrote", out.name)


def render_dock_scan() -> None:
    w, h = 1100, 720
    bg = Image.new("RGB", (w, h), (58, 58, 56))
    d = ImageDraw.Draw(bg)
    # off-white sticker, slightly rotated look via paste box
    d.rounded_rectangle([70, 50, 1030, 670], radius=8, fill=(247, 241, 228), outline=(90, 90, 88), width=2)

    title_f = font(28, bold=True)
    body_f = font(22)
    small_f = font(16)

    x, y = 110, 90
    d.text((x, y), "ATLANTA WAREHOUSE — RECEIVING EXCEPTION TAG", font=title_f, fill=(132, 38, 32))
    lines = [
        (body_f, (30, 30, 28), "Tracking: 1ZMD00000082"),
        (body_f, (30, 30, 28), "Ticket: OFF-00082 | Asset tag on carton: MD-00082"),
        (body_f, (132, 38, 32), "Serial scanned at dock: MISMATCH-0082"),
        (body_f, (30, 30, 28), "Register serial expected: MD-MO-050082"),
        (body_f, (30, 30, 28), "Status: HOLD — do not clear custody"),
        (small_f, (90, 88, 82), "Photo lead only. Not a receiving clearance."),
        (small_f, (90, 88, 82), "Atlanta inbound dock  ·  07 Jul 2026  ·  scanned by night receiving"),
    ]
    y = 155
    for fnt, col, text in lines:
        d.text((x, y), text, font=fnt, fill=col)
        y += 48 if fnt != small_f else 36

    # faux barcode
    by = 560
    d.rectangle([110, by, 990, by + 70], fill=(247, 241, 228))
    import random

    rng = random.Random(82)
    xbar = 120
    while xbar < 980:
        bw = rng.choice([2, 3, 4, 6, 8])
        if rng.random() > 0.35:
            d.rectangle([xbar, by, xbar + bw, by + 70], fill=(20, 20, 18))
        xbar += bw + rng.choice([1, 2, 3])

    out = ROOT / "receiving_exception_scan_1ZMD00000082.png"
    bg.save(out, "PNG")
    print("wrote", out.name)


def fix_hr_names() -> None:
    path = ROOT / "hr_employee_status.csv"
    rows = list(csv.DictReader(path.open(newline="")))
    if len(NAMES) < len(rows):
        raise SystemExit(f"need {len(rows)} names, have {len(NAMES)}")
    for i, row in enumerate(rows):
        row["employee_name"] = NAMES[i]
    with path.open("w", newline="") as f:
        w = csv.DictWriter(f, fieldnames=list(rows[0].keys()))
        w.writeheader()
        w.writerows(rows)
    print(f"renamed {len(rows)} employees")


def jitter_po_dates() -> None:
    path = ROOT / "hardware_purchase_orders.csv"
    rows = list(csv.DictReader(path.open(newline="")))
    days = [3, 7, 9, 12, 14, 18, 21, 22, 26, 27, 8, 16, 4, 11, 19, 23, 6, 13, 17, 28, 5, 15]
    for i, row in enumerate(rows):
        y, m, _ = (int(x) for x in row["order_date"].split("-"))
        d = days[i % len(days)]
        # keep valid calendar days (Feb)
        if m == 2 and d > 28:
            d = 19
        row["order_date"] = date(y, m, d).isoformat()
    with path.open("w", newline="") as f:
        w = csv.DictWriter(f, fieldnames=list(rows[0].keys()))
        w.writeheader()
        w.writerows(rows)
    print("jittered PO dates:", [r["order_date"] for r in rows[:6]], "...")


def rebuild_policy_pdf() -> None:
    styles = getSampleStyleSheet()
    title = ParagraphStyle("T", parent=styles["Heading1"], fontSize=16, spaceAfter=8)
    h = ParagraphStyle("H", parent=styles["Heading2"], fontSize=12, spaceBefore=10, spaceAfter=4)
    body = ParagraphStyle("B", parent=styles["Normal"], fontSize=9, leading=12, spaceAfter=4)
    small = ParagraphStyle("S", parent=styles["Normal"], fontSize=8, leading=10, textColor=colors.HexColor("#333333"))
    story = []
    story.append(Paragraph("IT Asset Management Policy", title))
    story.append(Paragraph("Yanou & Partners IT Services | Policy ITAM-001 | Version 3.2 | Effective 2026-01-01", small))
    story.append(Spacer(1, 0.12 * inch))
    sections = [
        (
            "1. Purpose and scope",
            "This policy establishes accountable custody, movement, return, loss, financial reconciliation, sanitization, "
            "and disposal controls for company-owned laptops, monitors, mobile devices, and network equipment.",
        ),
        (
            "2. System of record",
            "The IT asset register is the operational system of record for assignment and location. The fixed-asset ledger "
            "is the financial system of record for capitalized cost and net book value. Neither file may be updated from verbal claims alone.",
        ),
        (
            "3. Evidence precedence (conflicts)",
            "When sources disagree, apply this order and document the conflict: (1) fixed-asset ledger rows, verified disposal "
            "certificates, and carrier acceptance/delivery/receiving-scan triples; (2) approved transfers with non-blank approval IDs "
            "and received dates; (3) HR employment status; (4) service-desk ticket status fields; (5) technician notes and dock/exception "
            "images. Lower-precedence sources are leads only and never override higher-precedence transaction evidence.",
        ),
        (
            "4. Assignment and transfer",
            "Assets move only with a transfer ID, non-blank approval ID, and receiving acknowledgment. Blank approval IDs are control "
            "exceptions. Closed-office locations are not valid verified locations after consolidation.",
        ),
        (
            "5. Offboarding and returns",
            "Separated employees must return assigned assets within 10 business days. A carrier label is not proof of return. "
            "Return completion requires carrier acceptance, delivery confirmation, and a receiving scan that matches asset tag and "
            "serial number. Serial mismatches remain unresolved custody even if a label exists.",
        ),
        (
            "6. Receiving exceptions",
            "Shipment serial mismatches, damaged parcels, and dock exception photos are investigative leads. An exception image does "
            "not clear custody or mark an asset Available.",
        ),
        (
            "7. Capitalization and ledger expectations",
            "Capitalize individual assets with acquisition cost of $2,500 or more. A purchase-order field marking capitalization_approved "
            "= Yes authorizes capital treatment only for qualifying lines at or above the per-asset threshold; it does not require or imply "
            "that under-threshold lines appear on the fixed-asset ledger. Absence of an under-threshold asset from the ledger is expected "
            "and is not a Critical inventory-to-ledger certification blocker. Absence of a qualifying (≥ $2,500) asset from the ledger is a "
            "Critical finance exception that blocks certification until capitalized or formally written off.",
        ),
        (
            "8. Disposal and media sanitization",
            "Retirement requires an approved disposal path and a verified disposal certificate matching tag/serial (certificate_verified "
            "= Yes). Sanitization should align with NIST SP 800-88 Rev. 2. Recycler transfers without verified certificates remain "
            "Retired/Pending Disposal operationally even if the ledger already shows Disposed.",
        ),
        (
            "9. Financial reconciliation",
            "Reconcile operational inventory to the fixed-asset ledger each quarter. Explain every remaining difference class. Acquisition "
            "cost mismatches that are systematic but below $500 per asset require Finance remediation and do not alone block certification. "
            "Critical blockers are unresolved custody of capital assets, overdue returns without receiving proof, shipment serial mismatches, "
            "missing disposal certificates for retired assets, and missing ledger rows for assets at or above the capitalization threshold.",
        ),
        (
            "10. Certification",
            "The IT Operations Manager certifies the quarterly inventory only when Critical blockers are cleared or formally accepted by "
            "Finance and Internal Audit. HR validates employment status used in custody decisions. Internal Audit may test any custody chain.",
        ),
        (
            "11. Evidence and retention",
            "Evidence must identify the asset tag, serial number where applicable, transaction ID, date, and source system. Technician notes "
            "and images are retained as leads with the reconciliation package. Reconciliation packets are retained for seven years or longer "
            "when Finance or Internal Audit requires extended retention for capital assets.",
        ),
        (
            "12. Procurement and receiving",
            "Purchasing must reference asset tags on capital orders. Central Receiving records inbound scans that match tag and serial before "
            "custody is assigned. Partial shipments and carrier exceptions stay in a held status until receiving reconciles the inbound file "
            "to the purchase order and the asset register.",
        ),
        (
            "13. Physical inventory and audit support",
            "IT Operations runs cycle counts at major sites each quarter and documents variances as exceptions with owner, deadline, and "
            "financial exposure. Internal Audit may request custody chains for any capital asset; chains must cite transaction IDs rather "
            "than informal notes alone.",
        ),
    ]
    for heading, text in sections:
        story.append(Paragraph(heading, h))
        story.append(Paragraph(text, body))
    story.append(PageBreak())
    story.append(Paragraph("Appendix A — Control Matrix (figure)", h))
    story.append(Paragraph("Read with the narrative controls above. Severity for certification follows §7–§9.", body))
    img = ROOT / "ITAM_control_matrix.png"
    if img.exists():
        story.append(Spacer(1, 0.1 * inch))
        story.append(RLImage(str(img), width=6.5 * inch, height=3.6 * inch))
    story.append(Paragraph("Appendix B — Public Benchmark Sources", h))
    story.append(
        Preformatted(
            "NIST SP 800-53 Rev. 5 — Security and Privacy Controls\n"
            "NIST SP 800-88 Rev. 2 — Media Sanitization\n"
            "IRS Publication 946 — How To Depreciate Property",
            small,
        )
    )
    SimpleDocTemplate(str(ROOT / "IT_asset_management_policy.pdf"), pagesize=letter, title="ITAM-001").build(story)
    print("rebuilt policy PDF")


def rebuild_zip() -> None:
    inputs = [
        "it_asset_inventory.xlsx",
        "hr_employee_status.csv",
        "equipment_transfer_log.csv",
        "service_desk_offboarding.csv",
        "device_return_shipments.csv",
        "hardware_purchase_orders.csv",
        "fixed_asset_ledger.xlsx",
        "asset_disposal_records.csv",
        "regional_IT_notes.docx",
        "IT_asset_management_policy.pdf",
        "ITAM_control_matrix.png",
        "receiving_exception_scan_1ZMD00000082.png",
    ]
    zpath = ROOT / "Meridian_IT_Asset_Inputs.zip"
    with zipfile.ZipFile(zpath, "w", zipfile.ZIP_DEFLATED) as zf:
        for n in inputs:
            zf.write(ROOT / n, n)
    print("rebuilt", zpath.name, "with", len(inputs), "files")


def main() -> None:
    render_control_matrix()
    render_dock_scan()
    fix_hr_names()
    jitter_po_dates()
    rebuild_policy_pdf()
    rebuild_zip()


if __name__ == "__main__":
    main()
