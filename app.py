import streamlit as st
import zipfile
import json
import os
import io
import re
from datetime import datetime
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

st.set_page_config(page_title="Brandlight Payroll → חשבשבת", page_icon="📊", layout="wide")

# ============================================================
# CONFIG
# ============================================================
DEFAULT_ACCOUNTS = {
    "Gross Wages": "800000",
    "Expense Reimbursement": "800001",
    "Employer Federal & State Taxes": "800002",
    "Workers Compensation": "800003",
    "Employee Benefits": "800004",
    "Administrative Fee": "800005",
    "401(k) ER Contribution": "800006",
    "401(k) Establishment Fee": "800007",
}
DEFAULT_CREDIT = "540001"
RANGE_NAME = "SALARIES"

INVOICE_FIELDS_ORDER = [
    "Gross Wages",
    "Expense Reimbursement",
    "Employer Federal & State Taxes",
    "Workers' Compensation",
    "Employee Benefits",
    "Administrative Fee",
    "Other: 401(k) ER Contribution",
    "Other: 401(k) Establishment Fee",
]

FIELD_TO_ACCOUNT_NAME = {
    "Gross Wages": "Gross Wages",
    "Expense Reimbursement": "Expense Reimbursement",
    "Employer Federal & State Taxes": "Employer Federal & State Taxes",
    "Workers' Compensation": "Workers Compensation",
    "Employee Benefits": "Employee Benefits",
    "Administrative Fee": "Administrative Fee",
    "Other: 401(k) ER Contribution": "401(k) ER Contribution",
    "Other: 401(k) Establishment Fee": "401(k) Establishment Fee",
}

HISTORY_FILE = "payroll_history.json"

# ============================================================
# HISTORY MANAGEMENT
# ============================================================
def load_history():
    if os.path.exists(HISTORY_FILE):
        with open(HISTORY_FILE, 'r') as f:
            return json.load(f)
    return []

def save_history(history):
    with open(HISTORY_FILE, 'w') as f:
        json.dump(history, f, indent=2, ensure_ascii=False)

def add_to_history(payroll_data_list, history):
    existing_invoices = {h['invoice_number'] for h in history}
    added = 0
    for p in payroll_data_list:
        if p['invoice_number'] not in existing_invoices:
            entry = {
                'pay_date': p.get('pay_date', ''),
                'pay_date_hebrew': p.get('pay_date_hebrew', ''),
                'invoice_number': p.get('invoice_number', ''),
                'invoice_total': p['invoice'].get('total', 0),
                'components': {k: v for k, v in (p['invoice'].get('components', {}) or {}).items()},
            }
            history.append(entry)
            added += 1
    history.sort(key=lambda x: x.get('pay_date', ''))
    return added

# ============================================================
# PDF EXTRACTION
# ============================================================
def extract_pdf_data(uploaded_file):
    data = {"employees": [], "invoice": {}, "pay_date": "", "invoice_number": "", "invoice_total": 0}

    file_bytes = uploaded_file.read()
    uploaded_file.seek(0)

    with zipfile.ZipFile(io.BytesIO(file_bytes), 'r') as z:
        txt_files = sorted(
            [f for f in z.namelist() if f.endswith('.txt')],
            key=lambda x: int(x.replace('.txt', ''))
        )
        all_text = {}
        for tf in txt_files:
            page_num = int(tf.replace('.txt', ''))
            all_text[page_num] = z.read(tf).decode('utf-8', errors='replace')

    # Extract invoice
    for page_num in sorted(all_text.keys()):
        text = all_text[page_num]
        if 'TOTAL INVOICE' in text or 'SUB-TOTAL' in text:
            data['invoice'] = parse_invoice_page(text)
            break

    # Extract pay date
    for page_num in sorted(all_text.keys()):
        text = all_text[page_num]
        m = re.search(r'Pay Date\s+(\d{2}/\d{2}/\d{4})', text)
        if m:
            data['pay_date'] = m.group(1)
            break

    # Extract invoice number
    for page_num in sorted(all_text.keys()):
        text = all_text[page_num]
        m = re.search(r'Invoice(?:\s+#?\s*|\s+No\s+)(\d{7})', text)
        if m:
            data['invoice_number'] = m.group(1)
            break
    if not data['invoice_number']:
        for page_num in sorted(all_text.keys()):
            text = all_text[page_num]
            m = re.search(r'Invoice\s+(\d{7})', text)
            if m:
                data['invoice_number'] = m.group(1)
                break

    return data


def parse_invoice_page(text):
    invoice = {}
    field_names = [
        "Gross Wages", "Expense Reimbursement", "Employer Federal & State Taxes",
        "Workers' Compensation", "Employee Benefits", "Administrative Fee",
        "Other: 401(k) ER Contribution", "Other: 401(k) Establishment Fee",
    ]

    lines = [l.strip() for l in text.replace('\r', '').split('\n') if l.strip()]

    # Strategy 1: fields and amounts on separate sequential lines
    found_fields = []
    amount_lines = []
    for i, line in enumerate(lines):
        for field in field_names:
            if line == field or line.startswith(field):
                found_fields.append((i, field))
        if re.match(r'^[\d,]+\.\d{2}$', line):
            amount_lines.append((i, float(line.replace(',', ''))))

    result = {f: None for f in field_names}
    total = None

    if found_fields and amount_lines:
        last_field_idx = found_fields[-1][0]
        relevant_amounts = [(idx, amt) for idx, amt in amount_lines if idx > last_field_idx]
        for i, (_, field) in enumerate(found_fields):
            if i < len(relevant_amounts):
                result[field] = relevant_amounts[i][1]

    # If strategy 1 failed, try strategy 2: field and amount on same line
    if all(v is None for v in result.values()):
        for line in lines:
            for field in field_names:
                if field in line and result[field] is None:
                    amounts = re.findall(r'[\d,]+\.\d{2}', line)
                    if amounts:
                        result[field] = float(amounts[-1].replace(',', ''))

    # Extract total
    for line in lines:
        if 'TOTAL INVOICE' in line:
            amounts = re.findall(r'[\d,]+\.\d{2}', line)
            if amounts:
                total = float(amounts[-1].replace(',', ''))
        elif 'SUB-TOTAL' in line and total is None:
            amounts = re.findall(r'[\d,]+\.\d{2}', line)
            if amounts:
                total = float(amounts[-1].replace(',', ''))

    invoice['components'] = result
    invoice['total'] = total
    return invoice


def convert_date_to_hebrew(us_date):
    parts = us_date.split('/')
    if len(parts) == 3:
        return f"{parts[1]}/{parts[0]}/{parts[2]}"
    return us_date


# ============================================================
# EXCEL GENERATION
# ============================================================
def generate_excel(payroll_data_list, accounts, credit_account, history=None):
    wb = Workbook()

    hf = Font(name='Arial', bold=True, size=11, color='FFFFFF')
    hfill = PatternFill('solid', fgColor='2F5496')
    nf = Font(name='Arial', size=10)
    bf = Font(name='Arial', bold=True, size=10)
    tfill = PatternFill('solid', fgColor='E2EFDA')
    mfmt = '#,##0.00'
    tb = Border(left=Side(style='thin'), right=Side(style='thin'),
                top=Side(style='thin'), bottom=Side(style='thin'))

    def sc(cell, font=None, fill=None, fmt=None, align=None):
        if font: cell.font = font
        if fill: cell.fill = fill
        if fmt: cell.number_format = fmt
        if align: cell.alignment = align
        cell.border = tb

    # ---- Sheet 1: פקודת יומן ----
    ws = wb.active
    ws.title = 'פקודת יומן'

    ws.merge_cells('A1:H1')
    period = ""
    if payroll_data_list:
        dates = [p['pay_date_hebrew'] for p in payroll_data_list]
        period = f"{dates[0]} - {dates[-1]}"
    ws['A1'] = f'BRANDLIGHT INC. - פקודת יומן שכר | טווח: {RANGE_NAME} | {period}'
    ws['A1'].font = Font(name='Arial', bold=True, size=13, color='2F5496')
    ws['A1'].alignment = Alignment(horizontal='center')

    headers = ['תאריך', 'חשבון חובה 1', 'חשבון זכות 1', 'חשבון זכות 2',
               'פרטים', 'אסמכתא', 'חובה מט"ח', 'זכות מט"ח']
    for c, h in enumerate(headers, 1):
        sc(ws.cell(row=3, column=c, value=h), font=hf, fill=hfill,
           align=Alignment(horizontal='center', wrap_text=True))

    row = 4
    for p in payroll_data_list:
        inv = p['invoice']
        comps = inv.get('components', {})
        for field in INVOICE_FIELDS_ORDER:
            amt = comps.get(field, 0) or 0
            if amt == 0:
                continue
            acct_name = FIELD_TO_ACCOUNT_NAME.get(field, field)
            acct_num = accounts.get(acct_name, "")

            sc(ws.cell(row=row, column=1, value=p['pay_date_hebrew']), font=nf, fmt='@')
            ws.cell(row=row, column=1).number_format = '@'
            sc(ws.cell(row=row, column=2, value=acct_num), font=nf, fmt='@')
            ws.cell(row=row, column=2).number_format = '@'
            ws.cell(row=row, column=2).alignment = Alignment(horizontal='center')
            sc(ws.cell(row=row, column=3, value=credit_account), font=nf, fmt='@')
            ws.cell(row=row, column=3).number_format = '@'
            ws.cell(row=row, column=3).alignment = Alignment(horizontal='center')
            sc(ws.cell(row=row, column=4, value=''), font=nf, fmt='@')
            sc(ws.cell(row=row, column=5, value=acct_name), font=nf)
            sc(ws.cell(row=row, column=6, value=p['invoice_number']), font=nf, fmt='@')
            ws.cell(row=row, column=6).number_format = '@'
            ws.cell(row=row, column=6).alignment = Alignment(horizontal='center')
            sc(ws.cell(row=row, column=7, value=amt), font=nf, fmt=mfmt)
            sc(ws.cell(row=row, column=8, value=amt), font=nf, fmt=mfmt)
            row += 1

    widths = [14, 16, 16, 16, 34, 14, 18, 18]
    for c, w in enumerate(widths, 1):
        ws.column_dimensions[get_column_letter(c)].width = w

    # ---- Sheet 2: סיכום תקופה נוכחית ----
    ws2 = wb.create_sheet('סיכום תקופה')
    ws2.merge_cells('A1:K1')
    ws2['A1'] = 'BRANDLIGHT INC. - Invoice Summary (Current Batch)'
    ws2['A1'].font = Font(name='Arial', bold=True, size=14, color='2F5496')
    ws2['A1'].alignment = Alignment(horizontal='center')

    sh = ['Pay Date', 'Invoice #', 'Gross Wages', 'Expense Reimb.', 'ER Fed&State Tax',
          'Workers Comp', 'Emp Benefits', 'Admin Fee', '401k ER Contrib', '401k Est Fee', 'Total Invoice']
    for c, h in enumerate(sh, 1):
        sc(ws2.cell(row=3, column=c, value=h), font=hf, fill=hfill,
           align=Alignment(horizontal='center', wrap_text=True))

    for i, p in enumerate(payroll_data_list):
        r = 4 + i
        inv = p['invoice']
        comps = inv.get('components', {})
        sc(ws2.cell(row=r, column=1, value=p['pay_date_hebrew']), font=nf)
        sc(ws2.cell(row=r, column=2, value=p['invoice_number']), font=nf)
        for ci, field in enumerate(INVOICE_FIELDS_ORDER):
            val = comps.get(field, 0) or 0
            sc(ws2.cell(row=r, column=3+ci, value=val), font=nf, fmt=mfmt)
        sc(ws2.cell(row=r, column=11, value=inv.get('total', 0)), font=bf, fmt=mfmt)

    tr = 4 + len(payroll_data_list)
    sc(ws2.cell(row=tr, column=1, value='TOTAL'), font=bf, fill=tfill)
    sc(ws2.cell(row=tr, column=2), fill=tfill)
    for c in range(3, 12):
        sc(ws2.cell(row=tr, column=c,
           value=f'=SUM({get_column_letter(c)}4:{get_column_letter(c)}{tr-1})'),
           font=bf, fill=tfill, fmt=mfmt)
    for c in range(1, 12):
        ws2.column_dimensions[get_column_letter(c)].width = 16

    # ---- Sheet 3: מצטבר YTD ----
    if history:
        ws3 = wb.create_sheet('מצטבר YTD')
        ws3.merge_cells('A1:K1')
        ws3['A1'] = f'BRANDLIGHT INC. - YTD Cumulative {datetime.now().year}'
        ws3['A1'].font = Font(name='Arial', bold=True, size=14, color='2F5496')
        ws3['A1'].alignment = Alignment(horizontal='center')

        for c, h in enumerate(sh, 1):
            sc(ws3.cell(row=3, column=c, value=h), font=hf, fill=hfill,
               align=Alignment(horizontal='center', wrap_text=True))

        for i, h in enumerate(history):
            r = 4 + i
            sc(ws3.cell(row=r, column=1, value=h.get('pay_date_hebrew', '')), font=nf)
            sc(ws3.cell(row=r, column=2, value=h.get('invoice_number', '')), font=nf)
            comps = h.get('components', {})
            for ci, field in enumerate(INVOICE_FIELDS_ORDER):
                val = comps.get(field, 0) or 0
                sc(ws3.cell(row=r, column=3+ci, value=val), font=nf, fmt=mfmt)
            sc(ws3.cell(row=r, column=11, value=h.get('invoice_total', 0)), font=bf, fmt=mfmt)

        tr3 = 4 + len(history)
        sc(ws3.cell(row=tr3, column=1, value='YTD TOTAL'), font=bf, fill=tfill)
        sc(ws3.cell(row=tr3, column=2), fill=tfill)
        for c in range(3, 12):
            sc(ws3.cell(row=tr3, column=c,
               value=f'=SUM({get_column_letter(c)}4:{get_column_letter(c)}{tr3-1})'),
               font=bf, fill=tfill, fmt=mfmt)
        for c in range(1, 12):
            ws3.column_dimensions[get_column_letter(c)].width = 16

    return wb


# ============================================================
# STREAMLIT UI
# ============================================================
st.title("📊 Brandlight Payroll → חשבשבת")
st.markdown("**העלי קבצי Payroll PDF → קבלי פקודת יומן מוכנה לייבוא**")

# Load history
if 'history' not in st.session_state:
    st.session_state.history = load_history()

# Sidebar
with st.sidebar:
    st.header("⚙️ חשבונות")
    accounts = {}
    for name, default in DEFAULT_ACCOUNTS.items():
        accounts[name] = st.text_input(name, value=default, key=f"a_{name}")
    st.divider()
    credit = st.text_input("חשבון זכות (ספק)", value=DEFAULT_CREDIT)

    st.divider()
    st.header("📚 היסטוריה")
    st.metric("חשבוניות YTD", len(st.session_state.history))
    if st.session_state.history:
        ytd_total = sum(h.get('invoice_total', 0) or 0 for h in st.session_state.history)
        st.metric("סה\"כ YTD", f"${ytd_total:,.2f}")
    if st.button("🗑️ נקה היסטוריה"):
        st.session_state.history = []
        save_history([])
        st.rerun()

st.divider()

# Upload
uploaded_files = st.file_uploader(
    "📁 גררי קבצי US Payroll Consolidated PDF",
    type=['pdf'], accept_multiple_files=True
)

if uploaded_files:
    st.info(f"📄 {len(uploaded_files)} קבצים נבחרו")

    all_data = []
    errors = []
    progress = st.progress(0)

    for i, f in enumerate(uploaded_files):
        progress.progress((i+1)/len(uploaded_files), text=f"מעבד: {f.name}")
        try:
            pdata = extract_pdf_data(f)
            if pdata['pay_date']:
                pdata['pay_date_hebrew'] = convert_date_to_hebrew(pdata['pay_date'])
            else:
                pdata['pay_date_hebrew'] = ""
            if pdata['invoice'] and pdata['invoice'].get('total'):
                all_data.append(pdata)
            else:
                errors.append(f"⚠️ {f.name}: לא נמצאו נתוני חשבונית")
        except Exception as e:
            errors.append(f"❌ {f.name}: {str(e)}")

    progress.empty()
    for err in errors:
        st.warning(err)

    if all_data:
        all_data.sort(key=lambda x: x.get('pay_date', ''))

        # Summary table
        st.subheader("📋 חשבוניות שזוהו")
        cols = st.columns([2, 1])

        table_data = []
        batch_total = 0
        for p in all_data:
            t = p['invoice'].get('total', 0) or 0
            batch_total += t
            comps = p['invoice'].get('components', {})
            comp_sum = sum(v for v in comps.values() if v)
            match = "✅" if abs(comp_sum - t) < 0.02 else "❌"
            table_data.append({
                "תאריך": p['pay_date_hebrew'],
                "חשבונית": p['invoice_number'],
                "סה\"כ": f"${t:,.2f}",
                "רכיבים": sum(1 for v in comps.values() if v and v > 0),
                "אימות": match,
            })

        with cols[0]:
            st.dataframe(table_data, use_container_width=True, hide_index=True)
        with cols[1]:
            st.metric("סה\"כ אצווה נוכחית", f"${batch_total:,.2f}")
            st.metric("מספר חשבוניות", len(all_data))

        # Journal preview
        st.subheader("📝 תצוגה מקדימה — פקודת יומן")
        journal_preview = []
        for p in all_data:
            comps = p['invoice'].get('components', {})
            for field in INVOICE_FIELDS_ORDER:
                amt = comps.get(field, 0) or 0
                if amt == 0:
                    continue
                an = FIELD_TO_ACCOUNT_NAME.get(field, field)
                journal_preview.append({
                    "תאריך": p['pay_date_hebrew'],
                    "חובה": accounts.get(an, ""),
                    "זכות": credit,
                    "פרטים": an,
                    "אסמכתא": p['invoice_number'],
                    "חובה מט\"ח": f"${amt:,.2f}",
                    "זכות מט\"ח": f"${amt:,.2f}",
                })
        st.dataframe(journal_preview, use_container_width=True, height=400, hide_index=True)

        total_debit = sum(
            (p['invoice'].get('components', {}).get(f, 0) or 0)
            for p in all_data for f in INVOICE_FIELDS_ORDER
        )
        if abs(total_debit - batch_total) < 0.02:
            st.success(f"✅ מאוזן: חובה = זכות = ${total_debit:,.2f}")
        else:
            st.error(f"❌ לא מאוזן! חובה ${total_debit:,.2f} ≠ זכות ${batch_total:,.2f}")

        # Save to history
        added = add_to_history(all_data, st.session_state.history)
        if added > 0:
            save_history(st.session_state.history)
            st.toast(f"נוספו {added} חשבוניות להיסטוריה")

        # Download
        st.subheader("📥 הורדה")
        wb = generate_excel(all_data, accounts, credit, st.session_state.history)
        output = io.BytesIO()
        wb.save(output)
        output.seek(0)

        c1, c2 = st.columns(2)
        with c1:
            st.download_button(
                "⬇️ הורד פקודת יומן + סיכום + YTD",
                data=output, type="primary",
                file_name=f"Brandlight_Payroll_{datetime.now().strftime('%Y%m%d')}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

else:
    st.markdown("""
    ### 📌 הוראות שימוש
    1. **העלי קבצי PDF** — US Payroll Consolidated
    2. **בדקי נתונים** — אימות אוטומטי של רכיבים מול סה"כ חשבונית
    3. **הורידי Excel** — פקודת יומן מוכנה לחשבשבת + סיכום YTD מצטבר

    | רכיב | חשבון חובה |
    |---|---|
    | Gross Wages | 800000 |
    | Expense Reimbursement | 800001 |
    | Employer Fed & State Taxes | 800002 |
    | Workers Compensation | 800003 |
    | Employee Benefits | 800004 |
    | Administrative Fee | 800005 |
    | 401(k) ER Contribution | 800006 |
    | 401(k) Establishment Fee | 800007 |
    | **חשבון זכות** | **540001** |
    """)

# YTD History tab
if st.session_state.history:
    with st.expander("📚 היסטוריית חשבוניות YTD", expanded=False):
        hist_data = []
        for h in st.session_state.history:
            hist_data.append({
                "תאריך": h.get('pay_date_hebrew', ''),
                "חשבונית": h.get('invoice_number', ''),
                "סה\"כ": f"${h.get('invoice_total', 0):,.2f}",
            })
        st.dataframe(hist_data, use_container_width=True, hide_index=True)
