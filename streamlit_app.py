import re
import io
import pdfplumber
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
import streamlit as st

st.set_page_config(
    page_title="PDF → Excel | งบทดลอง",
    page_icon="📊",
    layout="centered"
)

# ─── CSS ────────────────────────────────────────────────────────────
st.markdown("""
<style>
[data-testid="stAppViewContainer"] { background: #f0f4f8; }
.main-card {
    background: white;
    border-radius: 16px;
    padding: 32px 36px;
    box-shadow: 0 4px 24px rgba(0,0,0,0.08);
    margin-bottom: 24px;
}
.stat-row { display: flex; gap: 12px; margin: 16px 0; }
.stat-box {
    flex: 1; background: #f0f9ff; border: 1.5px solid #bae6fd;
    border-radius: 10px; padding: 12px; text-align: center;
}
.stat-label { font-size: 12px; color: #64748b; margin-bottom: 4px; }
.stat-value { font-size: 14px; font-weight: 700; color: #1F4E79; }
</style>
""", unsafe_allow_html=True)


# ─── Helpers ────────────────────────────────────────────────────────
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

def detect_company(pdf_bytes):
    try:
        with pdfplumber.open(io.BytesIO(pdf_bytes)) as pdf:
            text = pdf.pages[0].extract_text() or ''
            lines = [l.strip() for l in text.splitlines() if l.strip()]
            return lines[0] if lines else ''
    except:
        return ''

def convert(pdf_bytes, company_name):
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

    for cells, fill in [('A4:A5',''), ('B4:B5',''), ('C4:D4',''), ('E4:F4',''), ('G4:H4','')]:
        ws.merge_cells(cells)
    top_labels = ['เลขที่บัญชี','ชื่อบัญชี','ยอดยกมา','','ยอดสะสมประจำงวด','','ยอดยกไป','']
    sub_labels = ['','','เดบิต','เครดิต','เดบิต','เครดิต','เดบิต','เครดิต']
    for col, val in enumerate(top_labels, 1):
        c = ws.cell(row=4, column=col, value=val or None)
        c.font=white_f; c.fill=hdr_fill; c.alignment=c_align; c.border=t_border
    for col, val in enumerate(sub_labels, 1):
        c = ws.cell(row=5, column=col, value=val or None)
        c.font=white_f; c.fill=sub_fill; c.alignment=c_align; c.border=t_border
    ws.row_dimensions[4].height = 24
    ws.row_dimensions[5].height = 20

    totals = [0.0]*6
    for i, row in enumerate(all_rows):
        er   = 6 + i
        fill = alt_fill if i % 2 == 0 else PatternFill("solid", fgColor="FFFFFF")
        vals = [clean_text(row[0]), clean_text(row[1]),
                clean_num(row[2]), clean_num(row[3]),
                clean_num(row[4]), clean_num(row[5]),
                clean_num(row[6]), clean_num(row[7])]
        for col, val in enumerate(vals, 1):
            c = ws.cell(row=er, column=col, value=val)
            c.font=data_f; c.fill=fill; c.border=t_border
            c.alignment = c_align if col==1 else (l_align if col==2 else r_align)
            if col > 2 and val is not None:
                c.number_format = num_fmt
        for j, v in enumerate(vals[2:]):
            if v: totals[j] += v

    tr = 6 + len(all_rows)
    for col in range(1, 9):
        c = ws.cell(row=tr, column=col)
        c.fill=tot_fill; c.font=total_f
        c.border=Border(left=thin,right=thin,top=medium,bottom=medium)
        c.alignment = c_align if col==1 else r_align
    ws.cell(row=tr, column=1).value = 'รวม'
    for col_idx, total in zip([3,4,5,6,7,8], totals):
        ws.cell(row=tr, column=col_idx).value = total
        ws.cell(row=tr, column=col_idx).number_format = num_fmt
    ws.row_dimensions[tr].height = 22

    ws.column_dimensions['A'].width = 16
    ws.column_dimensions['B'].width = 55
    for c in 'CDEFGH':
        ws.column_dimensions[c].width = 18
    ws.freeze_panes = 'C6'
    ws.page_setup.orientation = 'landscape'
    ws.page_setup.fitToPage = True
    ws.page_setup.fitToWidth = 1

    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    return buf, len(all_rows), totals


# ─── UI ─────────────────────────────────────────────────────────────
st.markdown("## 📊 PDF → Excel")
st.markdown("แปลงงบทดลอง PDF เป็น Excel อัตโนมัติ — รองรับไฟล์จากโปรแกรมบัญชีทั่วไป")
st.divider()

uploaded = st.file_uploader(
    "อัปโหลดไฟล์ PDF งบทดลอง",
    type=['pdf'],
    help="รองรับ PDF ที่สร้างจากโปรแกรมบัญชี (ไม่ใช่ภาพสแกน)"
)

if uploaded:
    pdf_bytes = uploaded.read()
    company_name = detect_company(pdf_bytes)

    if company_name:
        st.info(f"🏢 ตรวจพบ: **{company_name}**")

    if st.button("🔄 แปลงเป็น Excel", type="primary", use_container_width=True):
        with st.spinner("กำลังอ่านตาราง PDF และสร้างไฟล์ Excel..."):
            try:
                buf, row_count, totals = convert(pdf_bytes, company_name)
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
st.caption("💡 รองรับ PDF งบทดลองที่มีตารางชัดเจน (searchable PDF) เท่านั้น")
