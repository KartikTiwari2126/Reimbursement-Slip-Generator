import streamlit as st
from docx import Document
from docx.shared import Inches, Pt
from io import BytesIO
from datetime import date
from decimal import Decimal, InvalidOperation

# -------------------------------------------------------------------
# CONSTANT LETTER DETAILS — edit these once if authorized details change.
# -------------------------------------------------------------------
SENDER_NAME = "Aadarsh Kumar"
SENDER_ID = "MS25A081"
SENDER_DEPARTMENT = "Department of Management Studies"
INSTITUTE = "Indian Institute of Technology Madras"
ADDRESS = "Chennai- 600036"
COMMITTEE = "Sports and Cultural Committee, DoMS"

RECIPIENT_LINES = [
    "The Head of the Department",
    "Department of Management Studies",
    "Indian Institute of Technology Madras",
    "Chennai- 600036",
]

# Default bank details. They are editable in the app for every letter.
DEFAULT_BANK_NAME = "HDFC Bank"
DEFAULT_BANK_ACCOUNT_NUMBER = "50100354506601"
DEFAULT_BANK_ACCOUNT_HOLDER = "AADARSH KUMAR"
DEFAULT_BANK_BRANCH = "KILPAUK"
DEFAULT_BANK_IFSC = "HDFC0000124"


def money(value: Decimal) -> str:
    if value == value.quantize(Decimal("1")):
        return f"Rs. {value:,.0f}/-"
    return f"Rs. {value:,.2f}"


def parse_amount(value: str, field_name: str) -> Decimal:
    try:
        cleaned = value.strip().replace(",", "").replace("₹", "")
        cleaned = cleaned.replace("Rs.", "").replace("Rs", "").strip()
        amount = Decimal(cleaned)
        if amount < 0:
            raise ValueError
        return amount
    except (InvalidOperation, ValueError, AttributeError):
        raise ValueError(f"Enter a valid non-negative amount for {field_name}.")


def set_document_style(doc: Document):
    section = doc.sections[0]
    section.top_margin = Inches(0.8)
    section.bottom_margin = Inches(0.8)
    section.left_margin = Inches(1.0)
    section.right_margin = Inches(1.0)
    normal = doc.styles["Normal"]
    normal.font.name = "Times New Roman"
    normal.font.size = Pt(12)
    normal.paragraph_format.space_after = Pt(6)


def add_line(doc: Document, text: str = "", bold: bool = False):
    p = doc.add_paragraph()
    run = p.add_run(text)
    run.bold = bold
    return p


def create_letter(payment_type, letter_date, event_name, purchase_date, items,
                  total_amount, approved_budget, vendor_name="", bank_details=None):
    doc = Document()
    set_document_style(doc)

    add_line(doc, "From:")
    add_line(doc, SENDER_NAME)
    add_line(doc, SENDER_DEPARTMENT)
    add_line(doc, INSTITUTE)
    add_line(doc, ADDRESS)
    add_line(doc, f"Date: {letter_date.strftime('%d-%m-%Y')}")
    doc.add_paragraph()

    add_line(doc, "To:")
    for line in RECIPIENT_LINES:
        add_line(doc, line)
    doc.add_paragraph()

    add_line(doc, "Respected Sir,")
    doc.add_paragraph()

    if payment_type == "Payment to Self":
        subject = f"Sub: Request for Reimbursement of Expenses for {event_name}"
        opening = (
            f"I request reimbursement for expenses I incurred for {event_name} "
            f"on {purchase_date.strftime('%d-%m-%Y')}. The expense details are:"
        )
    else:
        subject = f"Sub: Request for Vendor Payment Processing for {event_name}"
        opening = (
            f"I request direct payment to {vendor_name} for purchases made for "
            f"{event_name} on {purchase_date.strftime('%d-%m-%Y')}. The invoice details are:"
        )

    add_line(doc, subject, bold=True)
    doc.add_paragraph(opening)

    for description, amount in items:
        add_line(doc, f"•  {description} — {money(amount)}")

    doc.add_paragraph()
    add_line(doc, f"Total Amount: {money(total_amount)}")
    add_line(doc, f"Approved Budget: {money(approved_budget)}")
    doc.add_paragraph()

    if payment_type == "Payment to Self":
        details = bank_details or {}
        doc.add_paragraph("Please process the reimbursement to the following bank account:")
        add_line(doc, f"Bank: {details['bank_name']}")
        add_line(doc, f"Account Number: {details['account_number']}")
        add_line(doc, f"Account Holder: {details['account_holder']}")
        add_line(doc, f"Branch: {details['branch']}")
        add_line(doc, f"IFSC: {details['ifsc']}")
    else:
        doc.add_paragraph(f"Kindly process the payment directly to {vendor_name}.")

    doc.add_paragraph("The relevant invoice(s) and supporting bills are attached for verification.")
    doc.add_paragraph("Thanking you.")
    doc.add_paragraph()
    add_line(doc, "Yours sincerely,")
    doc.add_paragraph()
    add_line(doc, f"{SENDER_NAME} ({SENDER_ID})")
    add_line(doc, COMMITTEE)

    output = BytesIO()
    doc.save(output)
    output.seek(0)
    return output.getvalue()


st.set_page_config(page_title="SportsCom Reimbursement Generator", page_icon="📄", layout="centered")
st.title("SportsCom Reimbursement Generator")
st.caption("DoMS, IIT Madras · Generate a standardized reimbursement letter")

payment_type = st.radio(
    "Payment type",
    ["Payment to Self", "Payment to Vendor"],
    horizontal=True,
)

col1, col2 = st.columns(2)
with col1:
    event_name = st.text_input("Event / purpose", placeholder="e.g. Chak De DoMS")
with col2:
    letter_date = st.date_input("Letter date", value=date.today())

purchase_date = st.date_input("Purchase / expense date", value=date.today())

vendor_name = ""
bank_details = None

if payment_type == "Payment to Vendor":
    vendor_name = st.text_input("Vendor name", placeholder="e.g. Asna Garments")
else:
    with st.expander("Bank details for reimbursement", expanded=True):
        st.caption("Pre-filled with Aadarsh Kumar's details. Edit these for each reimbursement as needed.")
        b1, b2 = st.columns(2)
        with b1:
            bank_name = st.text_input("Bank name", value=DEFAULT_BANK_NAME)
            account_number = st.text_input("Account number", value=DEFAULT_BANK_ACCOUNT_NUMBER)
            account_holder = st.text_input("Account holder", value=DEFAULT_BANK_ACCOUNT_HOLDER)
        with b2:
            branch = st.text_input("Branch", value=DEFAULT_BANK_BRANCH)
            ifsc = st.text_input("IFSC", value=DEFAULT_BANK_IFSC)
        bank_details = {
            "bank_name": bank_name.strip(),
            "account_number": account_number.strip(),
            "account_holder": account_holder.strip(),
            "branch": branch.strip(),
            "ifsc": ifsc.strip(),
        }

st.subheader("Expense items")
st.caption("Add each item once. Enter the amount for that line item.")

if "expense_rows" not in st.session_state:
    st.session_state.expense_rows = [{"description": "", "amount": ""}]

for i, row in enumerate(st.session_state.expense_rows):
    c1, c2, c3 = st.columns([5, 2, 1])
    with c1:
        row["description"] = st.text_input(
            "Description" if i == 0 else f"Description {i + 1}",
            value=row["description"],
            key=f"description_{i}",
            placeholder="e.g. Jerseys",
        )
    with c2:
        row["amount"] = st.text_input(
            "Amount (INR)" if i == 0 else f"Amount {i + 1} (INR)",
            value=row["amount"],
            key=f"amount_{i}",
            placeholder="46970",
        )
    with c3:
        st.write("")
        st.write("")
        if len(st.session_state.expense_rows) > 1 and st.button("✕", key=f"remove_{i}"):
            st.session_state.expense_rows.pop(i)
            st.rerun()

if st.button("＋ Add another item"):
    st.session_state.expense_rows.append({"description": "", "amount": ""})
    st.rerun()

col1, col2 = st.columns(2)
with col1:
    total_input = st.text_input("Total amount (INR)", placeholder="46970")
with col2:
    budget_input = st.text_input("Approved budget (INR)", placeholder="83630")

st.info("Sender, recipient, signature and committee name are fixed. Bank details are editable and included only for Payment to Self.")

if st.button("Generate DOCX", type="primary", use_container_width=True):
    errors = []
    if not event_name.strip():
        errors.append("Enter the event or purpose.")
    if payment_type == "Payment to Vendor" and not vendor_name.strip():
        errors.append("Enter the vendor name.")

    if payment_type == "Payment to Self":
        for key, label in [
            ("bank_name", "bank name"),
            ("account_number", "account number"),
            ("account_holder", "account holder"),
            ("branch", "branch"),
            ("ifsc", "IFSC"),
        ]:
            if not bank_details.get(key):
                errors.append(f"Enter the {label}.")

    items = []
    for i, row in enumerate(st.session_state.expense_rows, start=1):
        description = row["description"].strip()
        amount_text = row["amount"].strip()
        if not description and not amount_text:
            continue
        if not description or not amount_text:
            errors.append(f"Complete both description and amount for item {i}.")
            continue
        try:
            items.append((description, parse_amount(amount_text, f"item {i}")))
        except ValueError as exc:
            errors.append(str(exc))

    if not items:
        errors.append("Add at least one complete expense item.")

    try:
        total_amount = parse_amount(total_input, "total amount")
    except ValueError as exc:
        errors.append(str(exc))
        total_amount = Decimal("0")

    try:
        approved_budget = parse_amount(budget_input, "approved budget")
    except ValueError as exc:
        errors.append(str(exc))
        approved_budget = Decimal("0")

    if items and not errors:
        item_sum = sum((amount for _, amount in items), Decimal("0"))
        if item_sum != total_amount:
            errors.append(
                f"Item amounts add up to {money(item_sum)}, but Total amount is {money(total_amount)}. "
                "Please make them match."
            )

    if errors:
        for error in errors:
            st.error(error)
    else:
        file_bytes = create_letter(
            payment_type=payment_type,
            letter_date=letter_date,
            event_name=event_name.strip(),
            purchase_date=purchase_date,
            items=items,
            total_amount=total_amount,
            approved_budget=approved_budget,
            vendor_name=vendor_name.strip(),
            bank_details=bank_details,
        )
        safe_event = "".join(ch.lower() if ch.isalnum() else "_" for ch in event_name.strip()).strip("_")
        filename = f"{safe_event or 'event'}_{'self_reimbursement' if payment_type == 'Payment to Self' else 'vendor_payment'}.docx"
        st.success("Your letter is ready.")
        st.download_button(
            "Download Word document (.docx)",
            data=file_bytes,
            file_name=filename,
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            type="primary",
            use_container_width=True,
        )
