import streamlit as st
import io
import re
from datetime import datetime, timedelta
from openpyxl import load_workbook
from openpyxl.styles import PatternFill

st.set_page_config(page_title="Highlight Missing vs GL", layout="wide")
st.title("Highlight Missing Transactions vs GL")
st.write(
    "Upload the **Register** (bank transactions) and the **GL** (general ledger). "
    "Transactions in the register that are not found in the GL will be highlighted "
    "in orange. Checks are matched by **check number and amount**. "
    "Other transactions are matched by **amount and date** (within a 5-day window). "
    "Transfers ending in **7459** are ignored."
)

ORANGE_FILL = PatternFill(start_color="FFA500", end_color="FFA500", fill_type="solid")
TRANSFER_7459_KEYWORD = "7459"
DATE_TOLERANCE_DAYS = 5


def parse_date(value):
    """Parse a date value from either a datetime object or a string like '3/18/2026'."""
    if value is None:
        return None
    if isinstance(value, datetime):
        return value.date()
    s = str(value).strip()
    for fmt in ("%m/%d/%Y", "%Y-%m-%d", "%m/%d/%y"):
        try:
            return datetime.strptime(s, fmt).date()
        except ValueError:
            continue
    return None


def extract_gl_data(wb):
    """Extract transaction data from the GL file.

    GL layout (columns by index):
      I (9)  = Type, K (11) = Date, M (13) = Num,
      U (21) = Debit, W (23) = Credit
    In accounting for a cash/bank account:
      GL Debit  = money coming IN  (matches register credits)
      GL Credit = money going OUT  (matches register debits)

    Returns:
      gl_debits:  list of (date, amount) for non-check deposits/credits in
      gl_credits: list of (date, amount) for non-check payments out
      gl_checks:  list of (check_number_str, amount) for check payments
    """
    ws = wb.active
    gl_debits = []   # money in  -> matches register credits
    gl_credits = []  # money out -> matches register debits (non-checks)
    gl_checks = []   # (check_num_str, amount) -> matches register checks

    for row in ws.iter_rows(min_row=2, max_row=ws.max_row):
        date_val = row[10].value if len(row) > 10 else None   # Column K
        num_val = row[12].value if len(row) > 12 else None     # Column M
        debit_val = row[20].value if len(row) > 20 else None   # Column U
        credit_val = row[22].value if len(row) > 22 else None  # Column W

        if date_val is None or date_val == "Date":
            continue

        dt = parse_date(date_val)
        if dt is None:
            continue

        # Check if this GL row is a check (Num is a numeric check number)
        num_str = str(num_val).strip() if num_val is not None else ""
        is_gl_check = bool(re.match(r"^\d{4,6}$", num_str))

        if credit_val is not None:
            try:
                amt = round(float(credit_val), 2)
                if is_gl_check:
                    gl_checks.append((num_str, amt))
                else:
                    gl_credits.append((dt, amt))
            except (ValueError, TypeError):
                pass

        if debit_val is not None:
            try:
                gl_debits.append((dt, round(float(debit_val), 2)))
            except (ValueError, TypeError):
                pass

    return gl_debits, gl_credits, gl_checks


def extract_check_number(description):
    """Extract the check number from a register description like '© CHECK - 10442'.

    Returns the check number as a string, or None if not a check.
    """
    if description is None:
        return None
    desc = str(description)
    if "CHECK" not in desc.upper():
        return None
    m = re.search(r"(\d{4,6})", desc)
    return m.group(1) if m else None


def is_transfer_7459(description):
    """Check if a transaction description refers to a transfer ending in 7459."""
    if description is None:
        return False
    return TRANSFER_7459_KEYWORD in str(description)


def find_match(reg_date, reg_amount, pool, tolerance_days=DATE_TOLERANCE_DAYS):
    """Find and remove a matching (date, amount) entry from the pool.

    A match requires the same amount and a date within tolerance_days.
    If multiple matches exist, prefer the closest date.
    Returns True if a match was found and consumed.
    """
    best_idx = None
    best_diff = None

    for i, (gl_date, gl_amount) in enumerate(pool):
        if gl_amount == reg_amount:
            if reg_date is not None and gl_date is not None:
                diff = abs((reg_date - gl_date).days)
                if diff <= tolerance_days:
                    if best_diff is None or diff < best_diff:
                        best_idx = i
                        best_diff = diff
            elif reg_date is None or gl_date is None:
                # If either date is missing, match on amount alone
                if best_idx is None:
                    best_idx = i
                    best_diff = 0

    if best_idx is not None:
        pool.pop(best_idx)
        return True
    return False


def find_check_match(check_num, pool):
    """Find and remove a matching check number entry from the GL check pool.

    A match requires the same check number only (amounts may differ between
    the register and GL).
    Returns True if a match was found and consumed.
    """
    for i, (gl_check_num, gl_amount) in enumerate(pool):
        if gl_check_num == check_num:
            pool.pop(i)
            return True
    return False


def find_missing_and_highlight(register_wb, gl_debits, gl_credits, gl_checks):
    """Find register transactions missing from the GL and highlight them orange.

    Checks are matched by check number + amount against gl_checks.
    Non-check debits are matched by amount + date against gl_credits.
    Credits are matched by amount + date against gl_debits.

    Returns the modified workbook and counts of missing / total / ignored.
    """
    ws = register_wb.active

    # Build consumable pools
    gl_credit_pool = list(gl_credits)  # for matching non-check register debits (out)
    gl_debit_pool = list(gl_debits)    # for matching register credits (in)
    gl_check_pool = list(gl_checks)    # for matching register checks by number + amount

    total = 0
    missing = 0
    ignored = 0

    for row_idx in range(2, ws.max_row + 1):
        date_cell = ws.cell(row=row_idx, column=1).value
        if date_cell is None:
            continue
        date_str = str(date_cell)
        # Skip non-transaction rows
        if any(kw in date_str for kw in ["TOTALS", "Total", "Beginning", "Ending", "balance"]):
            continue

        reg_date = parse_date(date_cell)
        description = ws.cell(row=row_idx, column=3).value or ""
        debit_val = ws.cell(row=row_idx, column=4).value   # Debits (Out)
        credit_val = ws.cell(row=row_idx, column=5).value   # Credits (In)

        # Skip transfers ending in 7459
        if is_transfer_7459(description):
            ignored += 1
            continue

        total += 1
        is_missing = False
        check_num = extract_check_number(description)

        # Check debit (money out)
        if debit_val is not None:
            try:
                amt = round(float(debit_val), 2)
                if amt != 0:
                    if check_num:
                        # Checks: match by check number only
                        if not find_check_match(check_num, gl_check_pool):
                            is_missing = True
                    else:
                        # Non-checks: match by amount + date
                        if not find_match(reg_date, amt, gl_credit_pool):
                            is_missing = True
            except (ValueError, TypeError):
                pass

        # Check credit (money in): match by amount + date
        if credit_val is not None:
            try:
                amt = round(float(credit_val), 2)
                if amt != 0:
                    if not find_match(reg_date, amt, gl_debit_pool):
                        is_missing = True
            except (ValueError, TypeError):
                pass

        if is_missing:
            missing += 1
            # Highlight the entire row in orange
            for col_idx in range(1, ws.max_column + 1):
                ws.cell(row=row_idx, column=col_idx).fill = ORANGE_FILL

    return register_wb, total, missing, ignored


# --------------- Streamlit UI ---------------

col1, col2 = st.columns(2)
with col1:
    register_file = st.file_uploader("Upload Register (Excel)", type=["xlsx", "xls"])
with col2:
    gl_file = st.file_uploader("Upload GL (Excel)", type=["xlsx", "xls"])

if register_file and gl_file:
    if st.button("Compare & Highlight Missing", type="primary"):
        with st.spinner("Loading GL..."):
            gl_wb = load_workbook(io.BytesIO(gl_file.read()), data_only=True)
            gl_debits, gl_credits, gl_checks = extract_gl_data(gl_wb)

        st.info(
            f"GL loaded: **{len(gl_debits)}** deposit/debit entries, "
            f"**{len(gl_credits)}** non-check credit entries, "
            f"**{len(gl_checks)}** check entries."
        )

        with st.spinner("Comparing register against GL..."):
            register_wb = load_workbook(io.BytesIO(register_file.read()))
            register_wb, total, missing, ignored = find_missing_and_highlight(
                register_wb, gl_debits, gl_credits, gl_checks
            )

        # Summary
        matched = total - missing
        st.success(
            f"Done! **{total}** transactions checked, **{matched}** matched, "
            f"**{missing}** missing (highlighted orange), **{ignored}** transfers to 7459 ignored."
        )

        if missing > 0:
            st.warning(f"{missing} transaction(s) highlighted in orange are missing from the GL.")

        # Save to bytes for download
        output = io.BytesIO()
        register_wb.save(output)
        output.seek(0)

        st.download_button(
            label="Download Highlighted Register",
            data=output.getvalue(),
            file_name="register_highlighted.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.spreadsheetml",
            type="primary",
        )
