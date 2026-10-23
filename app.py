
import streamlit as st
from docx import Document
from docx.shared import Inches, Pt
from io import BytesIO
from datetime import date
from decimal import Decimal, InvalidOperation

# ============================================================
# FIXED SENDER AND RECEIVER DETAILS
# ============================================================

SENDER_NAME = "Kartikeya Tiwari"
SENDER_ID = "MS26A028"
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

# Default bank details for Payment to Self
DEFAULT_BANK_NAME = "HDFC Bank"
DEFAULT_ACCOUNT_NUMBER = "50100354506601"
DEFAULT_ACCOUNT_HOLDER = "AADARSH KUMAR"
DEFAULT_BRANCH = "KILPAUK"
DEFAULT_IFSC = "HDFC0000124"


# ============================================================
# HELPERS
# ============================================================

def parse_amount(value, field_name="amount"):
    try:
        cleaned = str(value).strip()
        cleaned = cleaned.replace(",", "").replace("₹", "")
        cleaned = cleaned.replace("Rs.", "").replace("Rs", "").strip()

        amount = Decimal(cleaned)

        if not amount.is_finite() or amount < 0:
            raise ValueError

        return amount

    except (InvalidOperation, ValueError, AttributeError):
        raise ValueError(
            f"Enter a valid non-negative number for {field_name}."
        )


def display_number(value):
    """Display 570 as 570, never as 5.7E+2."""
    rendered = format(value, "f")

    if "." in rendered:
        rendered = rendered.rstrip("0").rstrip(".")

    return rendered or "0"


def money(value):
    if value == value.quantize(Decimal("1")):
        return f"Rs. {value:,.0f}/-"

    return f"Rs. {value:,.2f}"


def calculate_total_callback():
    """Recalculate total before Streamlit reruns the app."""
    total = Decimal("0")

    rows = st.session_state.get("expense_rows", [])

    for i, row in enumerate(rows):
        raw = st.session_state.get(
            f"amount_{i}",
            row.get("amount", "")
        )

        if str(raw).strip():
            try:
                total += parse_amount(raw, f"item {i + 1}")
            except ValueError:
                # Do not overwrite the total while an amount is invalid.
                return

    # This callback runs before the total widget is instantiated.
    st.session_state["total_value"] = display_number(total)
    st.session_state["total_manually_edited"] = False


def mark_total_manually_edited():
    st.session_state["total_manually_edited"] = True


def reset_total_callback():
    calculate_total_callback()


# ============================================================
# DYNAMIC WORD DOCUMENT FORMATTING
# ============================================================

def choose_layout(item_count, text_length, payment_type):
    """
    Estimate content length and adjust formatting to help
    keep the letter on one page.
    """

    estimated_lines = 23 + item_count + (text_length // 85)

    if payment_type == "Payment to Self":
        estimated_lines += 5

    if estimated_lines <= 29:
        return 12, 1.15, 0.55, 0.70

    elif estimated_lines <= 35:
        return 11, 1.15, 0.45, 0.60

    elif estimated_lines <= 41:
        return 11, 1, 0.40, 0.55

    else:
        return 10, 1, 0.30, 0.50


def set_document_style(
    doc,
    font_size,
    line_spacing,
    vertical_margin,
    horizontal_margin
):
    section = doc.sections[0]

    section.top_margin = Inches(vertical_margin)
    section.bottom_margin = Inches(vertical_margin)
    section.left_margin = Inches(horizontal_margin)
    section.right_margin = Inches(horizontal_margin)

    normal = doc.styles["Normal"]
    normal.font.name = "Times New Roman"
    normal.font.size = Pt(font_size)

    normal.paragraph_format.space_before = Pt(0)
    normal.paragraph_format.space_after = Pt(0)
    normal.paragraph_format.line_spacing = line_spacing


def add_line(doc, text="", bold=False, font_size=10, line_spacing=1.0):
    paragraph = doc.add_paragraph()

    paragraph.paragraph_format.space_before = Pt(0)
    paragraph.paragraph_format.space_after = Pt(0)
    paragraph.paragraph_format.line_spacing = line_spacing

    run = paragraph.add_run(text)
    run.bold = bold
    run.font.name = "Times New Roman"
    run.font.size = Pt(font_size)

    return paragraph


# ============================================================
# GENERATE LETTER
# ============================================================

def create_letter(
    payment_type,
    letter_date,
    event_name,
    purchase_date,
    items,
    total_amount,
    approved_budget,
    vendor_name="",
    payment_details=None,
):
    payment_details = payment_details or {}

    text_length = (
        sum(len(description) for description, _ in items)
        + len(event_name)
        + len(vendor_name)
    )

    font_size, line_spacing, vertical_margin, horizontal_margin = (
        choose_layout(
            len(items),
            text_length,
            payment_type
        )
    )

    doc = Document()

    set_document_style(
        doc,
        font_size,
        line_spacing,
        vertical_margin,
        horizontal_margin
    )

    def line(text="", bold=False):
        add_line(
            doc,
            text,
            bold,
            font_size,
            line_spacing
        )

    # Sender
    line("From:")
    line(SENDER_NAME)
    line(SENDER_DEPARTMENT)
    line(INSTITUTE)
    line(ADDRESS)
    line(f"Date: {letter_date.strftime('%d-%m-%Y')}")
    line()

    # Receiver
    line("To:")

    for address_line in RECIPIENT_LINES:
        line(address_line)

    line()
    line("Respected Sir,")
    line()

    # Subject and introduction
    if payment_type == "Payment to Self":

        subject = (
            "Sub: Request for Reimbursement of Expenses "
            f"for {event_name}"
        )

        introduction = (
            f"I request reimbursement for expenses I incurred "
            f"for {event_name} on "
            f"{purchase_date.strftime('%d-%m-%Y')}. "
            "The expense details are:"
        )

    else:

        subject = (
            "Sub: Request for Vendor Payment Processing "
            f"for {event_name}"
        )

        introduction = (
            f"I request direct payment to {vendor_name} "
            f"for purchases made for {event_name} on "
            f"{purchase_date.strftime('%d-%m-%Y')}. "
            "The invoice details are:"
        )

    line(subject, bold=True)
    line()
    line(introduction)
    line()

    # Itemized expenses
    for description, amount in items:
        line(f"•  {description} — {money(amount)}")

    line()
    line(f"Total Amount: {money(total_amount)}")
    line(f"Approved Budget: {money(approved_budget)}")
    line()

    # Payment details differ by payment type
    if payment_type == "Payment to Self":

        line(
            "Please process the reimbursement to the "
            "following bank account:"
        )

    else:

        line(
            f"Please process the payment directly to "
            f"{vendor_name} using the following bank details:"
        )

    line()
    line(f"Bank: {payment_details['bank_name']}")
    line(f"Account Number: {payment_details['account_number']}")
    line(f"Account Holder: {payment_details['account_holder']}")
    line(f"Branch: {payment_details['branch']}")
    line(f"IFSC: {payment_details['ifsc']}")
    line()

    line(
        "The relevant invoice(s) and supporting bills "
        "are attached for verification."
    )

    line()
    line("Thanking you.")
    line()
    line("Yours sincerely,")
    line(f"{SENDER_NAME} ({SENDER_ID})")
    line(COMMITTEE)

    output = BytesIO()
    doc.save(output)
    output.seek(0)

    return output.getvalue(), font_size, line_spacing


# ============================================================
# STREAMLIT INTERFACE
# ============================================================

st.set_page_config(
    page_title="SportsCom Reimbursement Generator",
    page_icon="📄",
    layout="centered"
)

st.title("SportsCom Reimbursement Generator")
st.caption("DoMS, IIT Madras")

payment_type = st.radio(
    "Payment type",
    ["Payment to Self", "Payment to Vendor"],
    horizontal=True
)

col1, col2 = st.columns(2)

with col1:
    event_name = st.text_input(
        "Event / purpose",
        placeholder="e.g. Chak De DoMS"
    )

with col2:
    letter_date = st.date_input(
        "Letter date",
        value=date.today()
    )

purchase_date = st.date_input(
    "Purchase / expense date",
    value=date.today()
)


# ============================================================
# BANK DETAILS
# ============================================================

payment_details = {}
vendor_name = ""

if payment_type == "Payment to Self":

    with st.expander(
        "Bank details for reimbursement",
        expanded=True
    ):
        st.caption(
            "Default details are pre-filled. Edit if required."
        )

        col1, col2 = st.columns(2)

        with col1:
            payment_details["bank_name"] = st.text_input(
                "Bank name",
                value=DEFAULT_BANK_NAME,
                key="self_bank_name"
            )

            payment_details["account_number"] = st.text_input(
                "Account number",
                value=DEFAULT_ACCOUNT_NUMBER,
                key="self_account_number"
            )

            payment_details["account_holder"] = st.text_input(
                "Account holder",
                value=DEFAULT_ACCOUNT_HOLDER,
                key="self_account_holder"
            )

        with col2:
            payment_details["branch"] = st.text_input(
                "Branch",
                value=DEFAULT_BRANCH,
                key="self_branch"
            )

            payment_details["ifsc"] = st.text_input(
                "IFSC",
                value=DEFAULT_IFSC,
                key="self_ifsc"
            )

else:

    vendor_name = st.text_input(
        "Vendor name",
        placeholder="e.g. Asna Garments"
    )

    with st.expander(
        "Vendor bank details",
        expanded=True
    ):
        st.caption(
            "Enter the bank details supplied by this vendor."
        )

        col1, col2 = st.columns(2)

        with col1:
            payment_details["bank_name"] = st.text_input(
                "Vendor bank name",
                key="vendor_bank_name",
                placeholder="e.g. HDFC Bank"
            )

            payment_details["account_number"] = st.text_input(
                "Vendor account number",
                key="vendor_account_number",
                placeholder="Account number"
            )

            payment_details["account_holder"] = st.text_input(
                "Account holder / business name",
                key="vendor_account_holder",
                placeholder="Name on account"
            )

        with col2:
            payment_details["branch"] = st.text_input(
                "Vendor branch",
                key="vendor_branch",
                placeholder="Branch"
            )

            payment_details["ifsc"] = st.text_input(
                "Vendor IFSC",
                key="vendor_ifsc",
                placeholder="IFSC code"
            )


# ============================================================
# EXPENSE ITEMS
# ============================================================

st.subheader("Expense items")

st.caption(
    "The total updates when item amounts change. "
    "You can also edit it manually."
)

if "expense_rows" not in st.session_state:
    st.session_state.expense_rows = [
        {"description": "", "amount": ""}
    ]

if "total_value" not in st.session_state:
    st.session_state.total_value = "0"

if "total_manually_edited" not in st.session_state:
    st.session_state.total_manually_edited = False


for i, row in enumerate(st.session_state.expense_rows):

    col1, col2, col3 = st.columns([5, 2, 1])

    with col1:
        row["description"] = st.text_input(
            "Description" if i == 0 else f"Description {i + 1}",
            value=row["description"],
            key=f"description_{i}",
            placeholder="e.g. Jerseys"
        )

    with col2:
        row["amount"] = st.text_input(
            "Amount (INR)" if i == 0 else f"Amount {i + 1} (INR)",
            value=row["amount"],
            key=f"amount_{i}",
            placeholder="500",
            on_change=calculate_total_callback
        )

    with col3:
        st.write("")
        st.write("")

        if (
            len(st.session_state.expense_rows) > 1
            and st.button("✕", key=f"remove_{i}")
        ):
            st.session_state.expense_rows.pop(i)
            calculate_total_callback()
            st.rerun()


if st.button("＋ Add another item"):
    st.session_state.expense_rows.append(
        {"description": "", "amount": ""}
    )
    st.rerun()


# ============================================================
# TOTAL AMOUNT
# ============================================================

# Before creating the widget, synchronize the total from item values
# unless the user has intentionally edited the total.
if not st.session_state.total_manually_edited:

    running_total = Decimal("0")
    valid_total = True

    for i, row in enumerate(st.session_state.expense_rows):

        raw = st.session_state.get(
            f"amount_{i}",
            row.get("amount", "")
        )

        if str(raw).strip():
            try:
                running_total += parse_amount(
                    raw,
                    f"item {i + 1}"
                )
            except ValueError:
                valid_total = False
                break

    if valid_total:
        # This occurs before the total widget is instantiated.
        st.session_state.total_value = display_number(running_total)


total_label = (
    "Total spend (INR)"
    if payment_type == "Payment to Self"
    else "Total invoice amount (INR)"
)

st.text_input(
    total_label,
    key="total_value",
    on_change=mark_total_manually_edited,
    help="Editable. Use the recalculate button to restore the sum of item amounts."
)

st.button(
    "Recalculate total from items",
    on_click=reset_total_callback
)

budget_input = st.text_input(
    "Approved budget (INR)",
    placeholder="e.g. 83630"
)


# ============================================================
# GENERATE DOCX
# ============================================================

if st.button(
    "Generate DOCX",
    type="primary",
    use_container_width=True
):

    errors = []

    if not event_name.strip():
        errors.append("Enter the event or purpose.")

    if (
        payment_type == "Payment to Vendor"
        and not vendor_name.strip()
    ):
        errors.append("Enter the vendor name.")

    required_fields = [
        ("bank_name", "bank name"),
        ("account_number", "account number"),
        ("account_holder", "account holder"),
        ("branch", "branch"),
        ("ifsc", "IFSC"),
    ]

    for key, label in required_fields:
        if not payment_details.get(key, "").strip():

            prefix = (
                "Vendor "
                if payment_type == "Payment to Vendor"
                else ""
            )

            errors.append(f"Enter the {prefix}{label}.")

    items = []

    for i, row in enumerate(st.session_state.expense_rows):

        description = st.session_state.get(
            f"description_{i}",
            row.get("description", "")
        ).strip()

        raw_amount = st.session_state.get(
            f"amount_{i}",
            row.get("amount", "")
        ).strip()

        if not description and not raw_amount:
            continue

        if not description or not raw_amount:
            errors.append(
                f"Complete both description and amount for item {i + 1}."
            )
            continue

        try:
            amount = parse_amount(
                raw_amount,
                f"item {i + 1}"
            )
            items.append((description, amount))

        except ValueError as exc:
            errors.append(str(exc))

    if not items:
        errors.append("Add at least one complete expense item.")

    try:
        total_amount = parse_amount(
            st.session_state.total_value,
            "total amount"
        )
    except ValueError as exc:
        errors.append(str(exc))
        total_amount = Decimal("0")

    try:
        approved_budget = parse_amount(
            budget_input,
            "approved budget"
        )
    except ValueError as exc:
        errors.append(str(exc))
        approved_budget = Decimal("0")

    if errors:

        for error in errors:
            st.error(error)

    else:

        item_sum = sum(
            (amount for _, amount in items),
            Decimal("0")
        )

        if item_sum != total_amount:
            st.warning(
                f"Line items total {money(item_sum)}; "
                f"the manually entered total {money(total_amount)} "
                "will be used in the letter."
            )

        docx_bytes, font_size, line_spacing = create_letter(
            payment_type=payment_type,
            letter_date=letter_date,
            event_name=event_name.strip(),
            purchase_date=purchase_date,
            items=items,
            total_amount=total_amount,
            approved_budget=approved_budget,
            vendor_name=vendor_name.strip(),
            payment_details=payment_details
        )

        safe_event = "".join(
            ch.lower() if ch.isalnum() else "_"
            for ch in event_name.strip()
        ).strip("_")

        suffix = (
            "self_reimbursement"
            if payment_type == "Payment to Self"
            else "vendor_payment"
        )

        st.success(
            f"Letter generated. Formatting: Times New Roman, "
            f"{font_size} pt, line spacing {line_spacing:g}."
        )

        st.download_button(
            "Download Word document (.docx)",
            data=docx_bytes,
            file_name=f"{safe_event or 'event'}_{suffix}.docx",
            mime=(
                "application/vnd.openxmlformats-officedocument."
                "wordprocessingml.document"
            ),
            type="primary",
            use_container_width=True
        )
