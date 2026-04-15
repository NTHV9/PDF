import re
import io
import pdfplumber
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from collections import defaultdict
import streamlit as st

st.set_page_config(
    page_title="PDF → Excel Converter",
    page_icon="📊",
    layout="centered"
)

# ─── CSS (theme-aware — no hardcoded background) ────────────────────
st.markdown("""
<style>
.stat-label { font-size: 12px; margin-bottom: 4px; }
.stat-value { font-size: 14px; font-weight: 700; }
</style>
""", unsafe_allow_html=True)


# ─── Shared Helpers ──────────────────────────────────────────────────
def clean_num(val):
    if not val or str(val).strip() == '':
        return None
    try:
        return float(str(val).replace(',', '').strip())
    except:
        return None

def clean_text(val):
    if not val:
        return ''
    return re.sub(r'[\uf700-\uf7ff]', '', str(val)).strip()


# ─── PDF Type Detection ──────────────────────────────────────────────
def detect_pdf_type(pdf_bytes):
    """Returns 'statement' or 'trial_balance'."""
    try:
        with pdfplumber.open(io.BytesIO(pdf_bytes)) as pdf:
            text = (pdf.pages[0].extract_text() or '').upper()
            if 'STATEMENT OF ACCOUNT' in text:
                return 'statement'
    except:
        pass
    return 'trial_balance'


# ─── Trial Balance ───────────────────────────────────────────────────
def detect_company_tb(pdf_bytes):
    try:
        with pdfplumber.open(io.BytesIO(pdf_bytes)) as pdf:
            text = pdf.pages[0].extract_text() or ''
            lines = [l.strip() for l in text.splitlines() if l.strip()]
            return lines[0] if lines else ''
    except:
        return ''

def convert_trial_balance(pdf_bytes, company_name):
    all_rows = []
    with pdfplumber.open(io.BytesIO(pdf_bytes)) as pdf:
        for page in pdf.pages:
            for table in (page.extract_tables() or []):
                for row in table:
                    if not row or row[0] in ['เลขที่บัญชี', None]:
                        continue
                    if row[2] in ['ยอดยกมา', 'เดบิต']:
                        continue
                    if not row[0] or not str(row[0]).strip():
                        continue
                    all_rows.append(row)

    if not all_rows:
        raise ValueError("ไม่พบข้อมูลตารางใน PDF — กรุณาตรวจสอบว่าเป็น PDF งบทดลองที่รองรับ")

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "งบทดลอง"

    white_f  = Font(name='Arial', bold=True, size=10, color="FFFFFF")
    data_f   = Font(name='Arial', size=10)
    total_f  = Font(name='Arial', bold=True, size=10)
    c_align  = Alignment(horizontal='center', vertical='center', wrap_text=True)
    l_align  = Alignment(horizontal='left',   vertical='center')
    r_align  = Alignment(horizontal='right',  vertical='center')
    thin     = Side(style='thin',   color='BFBFBF')
    medium   = Side(style='medium', color='2E75B6')
    t_border = Border(left=thin, right=thin, top=thin, bottom=thin)
    hdr_fill = PatternFill("solid", fgColor="1F4E79")
    sub_fill = PatternFill("solid", fgColor="2E75B6")
    tot_fill = PatternFill("solid", fgColor="D6E4F0")
    alt_fill = PatternFill("solid", fgColor="F2F7FB")
    num_fmt  = '#,##0.00;[Red](#,##0.00);"-"'

    ws.merge_cells('A1:H1')
    ws['A1'] = company_name
    ws['A1'].font = Font(name='Arial', bold=True, size=14, color="1F4E79")
    ws['A1'].alignment = c_align

    ws.merge_cells('A2:H2')
    ws['A2'] = 'รายงานงบทดลอง'
    ws['A2'].font = Font(name='Arial', bold=True, size=12, color="2E75B6")
    ws['A2'].alignment = c_align

    ws.row_dimensions[1].height = 28
    ws.row_dimensions[2].height = 22
    ws.row_dimensions[3].height = 8

    for cells in ['A4:A5', 'B4:B5', 'C4:D4', 'E4:F4', 'G4:H4']:
        ws.merge_cells(cells)
    top_labels = ['เลขที่บัญชี', 'ชื่อบัญชี', 'ยอดยกมา', '', 'ยอดสะสมประจำงวด', '', 'ยอดยกไป', '']
    sub_labels = ['', '', 'เดบิต', 'เครดิต', 'เดบิต', 'เครดิต', 'เดบิต', 'เครดิต']
    for col, val in enumerate(top_labels, 1):
        c = ws.cell(row=4, column=col, value=val or None)
        c.font = white_f; c.fill = hdr_fill; c.alignment = c_align; c.border = t_border
    for col, val in enumerate(sub_labels, 1):
        c = ws.cell(row=5, column=col, value=val or None)
        c.font = white_f; c.fill = sub_fill; c.alignment = c_align; c.border = t_border
    ws.row_dimensions[4].height = 24
    ws.row_dimensions[5].height = 20

    totals = [0.0] * 6
    for i, row in enumerate(all_rows):
        er   = 6 + i
        fill = alt_fill if i % 2 == 0 else PatternFill("solid", fgColor="FFFFFF")
        vals = [clean_text(row[0]), clean_text(row[1]),
                clean_num(row[2]), clean_num(row[3]),
                clean_num(row[4]), clean_num(row[5]),
                clean_num(row[6]), clean_num(row[7])]
        for col, val in enumerate(vals, 1):
            c = ws.cell(row=er, column=col, value=val)
            c.font = data_f; c.fill = fill; c.border = t_border
            c.alignment = c_align if col == 1 else (l_align if col == 2 else r_align)
            if col > 2 and val is not None:
                c.number_format = num_fmt
        for j, v in enumerate(vals[2:]):
            if v: totals[j] += v

    tr = 6 + len(all_rows)
    for col in range(1, 9):
        c = ws.cell(row=tr, column=col)
        c.fill = tot_fill; c.font = total_f
        c.border = Border(left=thin, right=thin, top=medium, bottom=medium)
        c.alignment = c_align if col == 1 else r_align
    ws.cell(row=tr, column=1).value = 'รวม'
    for col_idx, total in zip([3, 4, 5, 6, 7, 8], totals):
        ws.cell(row=tr, column=col_idx).value = total
        ws.cell(row=tr, column=col_idx).number_format = num_fmt
    ws.row_dimensions[tr].height = 22

    ws.column_dimensions['A'].width = 16
    ws.column_dimensions['B'].width = 55
    for col_letter in 'CDEFGH':
        ws.column_dimensions[col_letter].width = 18
    ws.freeze_panes = 'C6'
    ws.page_setup.orientation = 'landscape'
    ws.page_setup.fitToPage = True
    ws.page_setup.fitToWidth = 1

    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    return buf, len(all_rows), totals


# ─── Statement of Account ────────────────────────────────────────────
def detect_company_soa(pdf_bytes):
    try:
        with pdfplumber.open(io.BytesIO(pdf_bytes)) as pdf:
            text = pdf.pages[0].extract_text() or ''
            lines = [l.strip() for l in text.splitlines() if l.strip()]
            name_parts = []
            for l in lines[:6]:
                if 'STATEMENT OF ACCOUNT' in l.upper(): continue
                if 'A/R' in l or 'Account No' in l:
                    part = l.split('A/R')[0].strip()
                    if part: name_parts.append(part)
                    continue
                if 'Print Date' in l or 'Page No' in l: break
                if len(l) < 40 and not any(kw in l for kw in ['Date', 'Folio', 'Debit', 'Credit']):
                    name_parts.append(l)
            return ' '.join(name_parts) if name_parts else ''
    except:
        return ''

_SOA_HEADER_COLS = ['Date', 'Folio', 'Description', 'Arrival', 'Departure', 'Voucher', 'Debit', 'Credit', 'Balance']

def _detect_soa_col_bounds(pdf_bytes):
    """Detect column boundaries dynamically from the header row AND actual data rows."""
    try:
        with pdfplumber.open(io.BytesIO(pdf_bytes)) as pdf:
            for page in pdf.pages[:2]:
                words = page.extract_words()
                rows_by_y = defaultdict(list)
                for w in words:
                    rows_by_y[round(w['top'] / 3) * 3].append(w)

                # ── Step 1: find header row to get approximate column x0 positions ──
                col = None
                header_y = None
                for y in sorted(rows_by_y.keys()):
                    row_words = rows_by_y[y]
                    texts = [w['text'] for w in row_words]
                    if 'Departure' in texts and 'Voucher' in texts:
                        c = {w['text']: w['x0'] for w in row_words if w['text'] in _SOA_HEADER_COLS}
                        if all(k in c for k in ['Date','Folio','Description','Arrival','Departure','Voucher','Debit']):
                            col = c
                            header_y = y
                            break

                if not col:
                    continue

                def mid(a, b): return (col[a] + col[b]) / 2

                # ── Step 2: scan data rows to find actual Arrival date x0 positions ──
                # Look for dd/mm/yy values between Description and Departure columns
                arr_data_x0s = []
                for y in sorted(rows_by_y.keys()):
                    if y <= header_y or y > 715:
                        continue
                    for w in rows_by_y[y]:
                        if (re.match(r'^\d{2}/\d{2}/\d{2}$', w['text']) and
                                col['Description'] < w['x0'] < col['Departure']):
                            arr_data_x0s.append(w['x0'])
                    if len(arr_data_x0s) >= 10:
                        break  # enough samples

                # DESC_MAX = 2px before the leftmost actual arrival date found
                # This is truly data-driven — adapts to each PDF automatically
                if arr_data_x0s:
                    desc_max = min(arr_data_x0s) - 2
                else:
                    desc_max = col['Arrival'] - 5  # fallback if no date data found

                return dict(
                    DATE_MAX      = mid('Date', 'Folio'),
                    FOLIO_MAX     = mid('Folio', 'Description'),
                    DESC_MAX      = desc_max,
                    ARR_MAX       = mid('Arrival', 'Departure'),
                    DEP_MAX       = mid('Departure', 'Voucher'),
                    VCH_MAX       = mid('Voucher', 'Debit'),
                    DEBIT_X1_MAX  = col.get('Credit', col['Debit'] + 50),
                    CREDIT_X1_MAX = col.get('Balance', col.get('Credit', col['Debit'] + 80) + 50),
                )
    except:
        pass
    # Fallback (Batch-style layout)
    return dict(DATE_MAX=74, FOLIO_MAX=133, DESC_MAX=218, ARR_MAX=298,
                DEP_MAX=356, VCH_MAX=411, DEBIT_X1_MAX=490, CREDIT_X1_MAX=546)

def extract_statement_rows(pdf_bytes):
    b = _detect_soa_col_bounds(pdf_bytes)
    DATE_MAX      = b['DATE_MAX']
    FOLIO_MAX     = b['FOLIO_MAX']
    DESC_MAX      = b['DESC_MAX']
    ARR_MAX       = b['ARR_MAX']
    DEP_MAX       = b['DEP_MAX']
    VCH_MAX       = b['VCH_MAX']
    DEBIT_X1_MAX  = b['DEBIT_X1_MAX']
    CREDIT_X1_MAX = b['CREDIT_X1_MAX']

    HEADER_WORDS = {'description', 'voucher', 'folio', 'arrival', 'departure',
                    'debit', 'credit', 'balance', 'date'}

    all_rows = []

    with pdfplumber.open(io.BytesIO(pdf_bytes)) as pdf:
        for page in pdf.pages:
            words = page.extract_words()
            rows_by_y = defaultdict(list)
            for w in words:
                rows_by_y[round(w['top'] / 3) * 3].append(w)

            # ── Pass 1: identify primary rows (date + folio) ──
            primary_ys   = []
            primary_data = {}

            for y in sorted(rows_by_y.keys()):
                if y < 195 or y > 715:
                    continue
                ws_words = sorted(rows_by_y[y], key=lambda w: w['x0'])
                date_w, folio_w, desc_w = [], [], []
                arr_w, dep_w, vch_w = [], [], []
                debit_w, credit_w, bal_w = [], [], []

                for w in ws_words:
                    x0, x1, t = w['x0'], w['x1'], w['text']
                    if x0 < DATE_MAX:          date_w.append(t)
                    elif x0 < FOLIO_MAX:       folio_w.append(t)
                    elif x0 < DESC_MAX:        desc_w.append(t)
                    elif x0 < ARR_MAX:         arr_w.append(t)
                    elif x0 < DEP_MAX:         dep_w.append(t)
                    elif x0 < VCH_MAX:         vch_w.append(t)
                    else:
                        if x1 <= DEBIT_X1_MAX:    debit_w.append(t)
                        elif x1 <= CREDIT_X1_MAX: credit_w.append(t)
                        else:                     bal_w.append(t)

                dv = ' '.join(date_w)
                fv = ' '.join(folio_w)
                if dv and fv and re.match(r'\d{2}/\d{2}/\d{2}', dv):
                    primary_ys.append(y)
                    primary_data[y] = {
                        'date': dv, 'folio': fv,
                        'desc': ' '.join(desc_w), 'arrival': ' '.join(arr_w),
                        'departure': ' '.join(dep_w), 'voucher': ' '.join(vch_w),
                        'debit': ' '.join(debit_w), 'credit': ' '.join(credit_w),
                        'balance': ' '.join(bal_w),
                    }

            if not primary_ys:
                continue

            # ── Pass 2: collect continuations with y positions ──
            # Use defaultdict keyed by nearest_py; each field stores [(y, text), ...]
            cont = defaultdict(lambda: {'desc': [], 'voucher': [],
                                        'debit': [], 'credit': [], 'balance': []})

            for y in sorted(rows_by_y.keys()):
                if y < 195 or y > 715:
                    continue
                if y in primary_data:
                    continue  # already handled as a primary row

                # Find nearest primary using midpoint rule
                nearest_py = None
                if y < primary_ys[0]:
                    # Before first primary — belongs to the first primary
                    nearest_py = primary_ys[0]
                else:
                    for i in range(len(primary_ys) - 1):
                        mid = (primary_ys[i] + primary_ys[i + 1]) / 2
                        if y <= mid:
                            nearest_py = primary_ys[i]
                            break
                    if nearest_py is None:
                        nearest_py = primary_ys[-1]

                # Skip if too far from nearest primary (aging/summary section)
                if abs(y - nearest_py) > 60:
                    continue

                ws_words = sorted(rows_by_y[y], key=lambda w: w['x0'])
                desc_w, vch_w, debit_w, credit_w, bal_w = [], [], [], [], []
                for w in ws_words:
                    x0, x1, t = w['x0'], w['x1'], w['text']
                    if FOLIO_MAX <= x0 < DESC_MAX:
                        desc_w.append(t)
                    elif DEP_MAX <= x0 < VCH_MAX:
                        vch_w.append(t)
                    elif x0 >= VCH_MAX:
                        if x1 <= DEBIT_X1_MAX:    debit_w.append(t)
                        elif x1 <= CREDIT_X1_MAX: credit_w.append(t)
                        else:                     bal_w.append(t)

                desc_cont   = ' '.join(desc_w).strip()
                vch_cont    = ' '.join(vch_w).strip()
                debit_cont  = ''.join(debit_w)
                credit_cont = ''.join(credit_w)
                bal_cont    = ''.join(bal_w)

                # Skip repeated header rows
                if desc_cont.lower() in HEADER_WORDS or vch_cont.lower() in HEADER_WORDS:
                    continue

                if desc_cont:   cont[nearest_py]['desc'].append((y, desc_cont))
                if vch_cont:    cont[nearest_py]['voucher'].append((y, vch_cont))
                if debit_cont:  cont[nearest_py]['debit'].append((y, debit_cont))
                if credit_cont: cont[nearest_py]['credit'].append((y, credit_cont))
                if bal_cont:    cont[nearest_py]['balance'].append((y, bal_cont))

            # ── Build final field values in correct y order ──
            for py in primary_ys:
                row = primary_data[py]

                # Voucher: sort all fragments (primary + continuations) by y,
                # then concatenate WITHOUT spaces — long references wrap mid-token
                vch_all = sorted(cont[py]['voucher'] + [(py, row['voucher'])],
                                 key=lambda x: x[0])
                row['voucher'] = ''.join(v for _, v in vch_all if v)

                # Desc: sort by y, join WITH spaces (name/text wraps with natural spaces)
                desc_all = sorted(cont[py]['desc'] + [(py, row['desc'])],
                                  key=lambda x: x[0])
                row['desc'] = ' '.join(d for _, d in desc_all if d)

                # Debit/Credit/Balance: sort by y, concatenate WITHOUT spaces
                # (digits split across lines, e.g. "151,130.0" + "0" → "151,130.00")
                for field in ('debit', 'credit', 'balance'):
                    parts = sorted(cont[py][field] + [(py, row[field])],
                                   key=lambda x: x[0])
                    row[field] = ''.join(v for _, v in parts if v)

                all_rows.append(row)

    return all_rows

def convert_statement(pdf_bytes, client_name, print_date):
    rows = extract_statement_rows(pdf_bytes)

    if not rows:
        raise ValueError("ไม่พบข้อมูลใน PDF — กรุณาตรวจสอบว่าเป็น Statement of Account ที่รองรับ")

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Statement of Account"

    white_f  = Font(name='Arial', bold=True, size=10, color="FFFFFF")
    total_f  = Font(name='Arial', bold=True, size=10)
    title_f  = Font(name='Arial', bold=True, size=13, color="1F4E79")
    sub_f    = Font(name='Arial', bold=True, size=11, color="2E75B6")
    c_align  = Alignment(horizontal='center', vertical='center', wrap_text=True)
    l_align  = Alignment(horizontal='left',   vertical='center', wrap_text=True)
    r_align  = Alignment(horizontal='right',  vertical='center')
    thin     = Side(style='thin',   color='BFBFBF')
    medium   = Side(style='medium', color='2E75B6')
    t_border = Border(left=thin, right=thin, top=thin, bottom=thin)
    hdr_fill = PatternFill("solid", fgColor="1F4E79")
    tot_fill = PatternFill("solid", fgColor="D6E4F0")
    alt_fill = PatternFill("solid", fgColor="F2F7FB")
    num_fmt  = '#,##0.00;[Red](#,##0.00);"-"'

    ws.merge_cells('A1:I1')
    ws['A1'] = "STATEMENT OF ACCOUNT"
    ws['A1'].font = title_f; ws['A1'].alignment = c_align
    ws.row_dimensions[1].height = 26

    ws.merge_cells('A2:G2')
    ws['A2'] = client_name
    ws['A2'].font = sub_f; ws['A2'].alignment = l_align
    ws['H2'] = "Print Date:"
    ws['H2'].font = Font(name='Arial', bold=True, size=9); ws['H2'].alignment = r_align
    ws['I2'] = print_date
    ws['I2'].font = Font(name='Arial', size=9); ws['I2'].alignment = c_align
    ws.row_dimensions[2].height = 20
    ws.row_dimensions[3].height = 5

    headers = ['Date', 'Folio', 'Description', 'Arrival', 'Departure', 'Voucher', 'Debit', 'Credit', 'Balance']
    for col, h in enumerate(headers, 1):
        c = ws.cell(row=4, column=col, value=h)
        c.font = white_f; c.fill = hdr_fill
        c.alignment = c_align; c.border = t_border
    ws.row_dimensions[4].height = 22

    for i, row in enumerate(rows):
        er = 5 + i
        fill = alt_fill if i % 2 == 0 else PatternFill("solid", fgColor="FFFFFF")
        vals = [
            row['date'], row['folio'], row['desc'],
            row['arrival'], row['departure'], row['voucher'],
            clean_num(row['debit']), clean_num(row['credit']), clean_num(row['balance'])
        ]
        for col, val in enumerate(vals, 1):
            c = ws.cell(row=er, column=col, value=val)
            c.font = Font(name='Arial', size=9)
            c.fill = fill; c.border = t_border
            if col == 3:        c.alignment = l_align
            elif col >= 7:
                c.alignment = r_align
                if val is not None: c.number_format = num_fmt
            else:               c.alignment = c_align
        ws.row_dimensions[er].height = 16

    tr = 5 + len(rows)
    td = sum(clean_num(r['debit'])   or 0 for r in rows)
    tc = sum(clean_num(r['credit'])  or 0 for r in rows)
    tb = sum(clean_num(r['balance']) or 0 for r in rows)

    ws.merge_cells(f'A{tr}:F{tr}')
    c = ws.cell(row=tr, column=1, value='รวมทั้งสิ้น / Grand Total')
    c.font = total_f; c.alignment = c_align
    c.fill = tot_fill; c.border = Border(left=thin, right=thin, top=medium, bottom=medium)
    for col in range(2, 7):
        c = ws.cell(row=tr, column=col)
        c.fill = tot_fill; c.border = Border(left=thin, right=thin, top=medium, bottom=medium)
    for col, val in [(7, td), (8, tc if tc else None), (9, tb)]:
        c = ws.cell(row=tr, column=col, value=val)
        c.fill = tot_fill; c.font = total_f
        c.border = Border(left=thin, right=thin, top=medium, bottom=medium)
        c.alignment = r_align
        if val is not None: c.number_format = num_fmt
    ws.row_dimensions[tr].height = 22

    col_widths = [12, 9, 42, 12, 12, 22, 16, 14, 16]
    for i, cw in enumerate(col_widths, 1):
        ws.column_dimensions[openpyxl.utils.get_column_letter(i)].width = cw

    ws.freeze_panes = 'A5'
    ws.page_setup.orientation = 'landscape'
    ws.page_setup.fitToPage = True
    ws.page_setup.fitToWidth = 1

    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    return buf, len(rows), td, tc, tb

def get_print_date(pdf_bytes):
    try:
        with pdfplumber.open(io.BytesIO(pdf_bytes)) as pdf:
            text = pdf.pages[0].extract_text() or ''
            m = re.search(r'Print\s+Date\s+(\d{2}/\d{2}/\d{2})', text)
            if m: return m.group(1)
    except:
        pass
    return ''


# ─── UI ─────────────────────────────────────────────────────────────
st.markdown("## 📊 PDF → Excel Converter")
st.markdown("รองรับ 2 ประเภท: **งบทดลอง** และ **Statement of Account**")
st.divider()

uploaded = st.file_uploader(
    "อัปโหลดไฟล์ PDF",
    type=['pdf'],
    help="รองรับ PDF ที่สร้างจากโปรแกรมบัญชี (ไม่ใช่ภาพสแกน)"
)

if uploaded:
    pdf_bytes = uploaded.read()
    pdf_type  = detect_pdf_type(pdf_bytes)

    if pdf_type == 'statement':
        client_name = detect_company_soa(pdf_bytes)
        print_date  = get_print_date(pdf_bytes)
        st.info(f"📄 ตรวจพบ: **Statement of Account** — {client_name}")

        if st.button("🔄 แปลงเป็น Excel", type="primary", use_container_width=True):
            with st.spinner("กำลังอ่าน PDF และสร้างไฟล์ Excel..."):
                try:
                    buf, row_count, td, tc, tb = convert_statement(pdf_bytes, client_name, print_date)
                    xlsx_name = uploaded.name.replace('.pdf', '.xlsx').replace('.PDF', '.xlsx')

                    st.success(f"✅ แปลงสำเร็จ — **{row_count:,} รายการ**")

                    col1, col2, col3 = st.columns(3)
                    def fmt(n): return f"{n:,.2f}" if n else "-"
                    col1.metric("Total Debit",   fmt(td))
                    col2.metric("Total Credit",  fmt(tc) if tc else "-")
                    col3.metric("Total Balance", fmt(tb))

                    st.download_button(
                        label="⬇️ ดาวน์โหลด Excel",
                        data=buf,
                        file_name=xlsx_name,
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        use_container_width=True,
                        type="primary"
                    )
                except Exception as e:
                    st.error(f"❌ เกิดข้อผิดพลาด: {e}")

    else:
        company_name = detect_company_tb(pdf_bytes)
        st.info(f"🏢 ตรวจพบ: **งบทดลอง** — {company_name}")

        if st.button("🔄 แปลงเป็น Excel", type="primary", use_container_width=True):
            with st.spinner("กำลังอ่านตาราง PDF และสร้างไฟล์ Excel..."):
                try:
                    buf, row_count, totals = convert_trial_balance(pdf_bytes, company_name)
                    xlsx_name = uploaded.name.replace('.pdf', '.xlsx').replace('.PDF', '.xlsx')

                    st.success(f"✅ แปลงสำเร็จ — **{row_count:,} รายการ**")

                    col1, col2, col3 = st.columns(3)
                    def fmt(n): return f"{n:,.2f}" if n else "-"
                    col1.metric("ยอดยกมา (Dr)",   fmt(totals[0]))
                    col2.metric("ยอดสะสม (Dr)",   fmt(totals[2]))
                    col3.metric("ยอดยกไป (Dr)",   fmt(totals[4]))

                    st.download_button(
                        label="⬇️ ดาวน์โหลด Excel",
                        data=buf,
                        file_name=xlsx_name,
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        use_container_width=True,
                        type="primary"
                    )
                except Exception as e:
                    st.error(f"❌ เกิดข้อผิดพลาด: {e}")

st.divider()
st.caption("💡 ระบบตรวจจับประเภท PDF อัตโนมัติ — รองรับเฉพาะ searchable PDF (ไม่ใช่ภาพสแกน)")
