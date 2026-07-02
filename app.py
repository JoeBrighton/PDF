import streamlit as st
import pdfplumber
import re
import tempfile
import os
import io
import difflib
from datetime import datetime
from collections import defaultdict
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils import get_column_letter

st.set_page_config(page_title="Clipboard Invoice Reconciliation", layout="wide")
st.title("Clipboard Invoice Reconciliation")
st.caption("Upload a Clipboard invoice PDF and an Empion punch report to generate the reconciliation.")
st.info(
    "🔒 **Privacy:** Uploaded files are processed in memory only and never stored. "
    "All data is discarded when you close the browser or upload new files. "
    "Nothing is saved to any server or database.",
    icon=None,
)

# ─────────────────────────────────────────────────────────────────────────────
# HELPERS
# ─────────────────────────────────────────────────────────────────────────────

def norm(n):
    return n.lower().replace("'", " ").replace("'", " ").replace("-", " ").strip()

def parse_dt(s):
    try:
        return datetime.strptime(s.strip(), "%m/%d/%Y %I:%M %p")
    except:
        return None

def norm_role(r):
    return "CMA" if "CERTIFIED" in r else r

# ─────────────────────────────────────────────────────────────────────────────
# INVOICE PARSER
# ─────────────────────────────────────────────────────────────────────────────

def parse_clipboard_invoice(pdf_bytes):
    with tempfile.NamedTemporaryFile(delete=False, suffix=".pdf") as tmp:
        tmp.write(pdf_bytes)
        tmp_path = tmp.name
    try:
        with pdfplumber.open(tmp_path) as pdf:
            text = "\n".join(page.extract_text() or "" for page in pdf.pages)
    finally:
        os.unlink(tmp_path)

    inv_num  = re.search(r'INVOICE\s+(\d+(?:-\d+)?)', text)
    facility = re.search(r'Bill To:\s*\n(.+)', text)
    period_s = re.search(r'Period Start Date:\s+(\d{2}/\d{2}/\d{4})', text)
    period_e = re.search(r'Period End Date:\s+(\d{2}/\d{2}/\d{4})', text)
    balance  = re.search(r'Balance Due:\s+\$([\d,]+\.\d{2})', text)

    meta = {
        'invoice':      inv_num.group(1) if inv_num else '',
        'facility':     facility.group(1).strip() if facility else '',
        'period_start': period_s.group(1) if period_s else '',
        'period_end':   period_e.group(1) if period_e else '',
        'balance_due':  float(balance.group(1).replace(',', '')) if balance else 0,
    }

    ROLES = r'(CNA|RN|LPN|LVN|RN/DON|CERTIFIED MEDICATION AIDE)'
    NORMAL_RE = re.compile(
        r'(\d{2}/\d{2}/\d{4}),\s+' + ROLES + r'\s+(.+?),\s+(AM|PM|NOC)\s+([\d.]+)\s+\$([\d.]+)\s+\$([\d,]+\.\d{2})'
    )
    LC_RE = re.compile(
        r'(\d{2}/\d{2}/\d{4}),\s+' + ROLES + r'\s+(.+?),\s+(AM|PM|NOC),\s+\(Late Cancel\)\s+([\d.]+)\s+\$([\d.]+)\s+\$([\d,]+\.\d{2})'
    )
    WB_RE = re.compile(
        r'(\d{2}/\d{2}/\d{4}),\s+' + ROLES + r'\s+(.+?),\s+(AM|PM|NOC)\s+\(Shift has an approved Worked Break Payment Request for\s+([\d.]+)\s+\$([\d.]+)\s+\$([\d,]+\.\d{2})'
    )
    START_RE = re.compile(r'Shift start:\s+(\d{2}/\d{2}/\d{4}\s+\d{1,2}:\d{2}\s+[AP]M)')
    END_RE   = re.compile(r'Shift end:\s+(\d{2}/\d{2}/\d{4}\s+\d{1,2}:\d{2}\s+[AP]M)')

    lines = text.split('\n')
    shifts = []
    for i, line in enumerate(lines):
        lc = wb = False
        m = LC_RE.search(line)
        if m:
            lc = True
        else:
            m = WB_RE.search(line)
            if m:
                wb = True
            else:
                m = NORMAL_RE.search(line)
        if not m:
            continue
        date    = datetime.strptime(m.group(1), "%m/%d/%Y").strftime("%m/%d")
        role    = norm_role(m.group(2))
        emp     = m.group(3).strip()
        shift   = m.group(4)
        inv_hrs = float(m.group(5))
        rate    = float(m.group(6))
        block   = ' '.join(lines[i:i+6])
        ms = START_RE.search(block)
        me = END_RE.search(block)
        shifts.append({
            'date': date, 'role': role, 'emp': emp, 'shift': shift,
            'lc': lc, 'worked_brk': wb, 'inv_hrs': inv_hrs, 'rate': rate,
            'start': parse_dt(ms.group(1)) if ms else None,
            'end':   parse_dt(me.group(1)) if me else None,
        })
    return meta, shifts

# ─────────────────────────────────────────────────────────────────────────────
# EMPION PARSER
# ─────────────────────────────────────────────────────────────────────────────

def parse_empion(xlsx_bytes):
    from openpyxl import load_workbook
    with tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx") as tmp:
        tmp.write(xlsx_bytes)
        tmp_path = tmp.name
    try:
        wb = load_workbook(tmp_path, data_only=True)
        ws = wb.active
        hdr = [c.value for c in next(ws.iter_rows(min_row=1, max_row=1))]
        name_col = hdr.index('Full Name')
        in_col   = hdr.index('In Time')
        out_col  = hdr.index('Out Time')
        hrs_col  = hdr.index('Total Hours')
        raw = []
        for row in ws.iter_rows(min_row=2, values_only=True):
            fn = row[name_col]; it = row[in_col]; ot = row[out_col]; rh = row[hrs_col]
            if not fn or not it: continue
            p = fn.split(', ', 1)
            fname = f"{p[1]} {p[0]}" if len(p) == 2 else fn
            raw.append({'name': fname, 'date': it.strftime("%m/%d"), 'in': it, 'out': ot, 'raw': rh})
    finally:
        os.unlink(tmp_path)

    combined = defaultdict(list)
    for r in raw:
        combined[(norm(r['name']), r['date'])].append(r)

    emp_idx = {}
    for key, entries in combined.items():
        entries.sort(key=lambda x: x['in'])
        total_raw = sum(e['raw'] for e in entries)
        deduct = int(total_raw // 8) * 0.5
        adj = round(total_raw - deduct, 2)
        emp_idx[key] = {
            'name': entries[0]['name'], 'date': entries[0]['date'],
            'in': entries[0]['in'], 'out': entries[-1]['out'],
            'raw': total_raw, 'deduct': deduct, 'adj': adj,
            'split': len(entries) > 1, 'entries': entries,
        }
    return emp_idx

# ─────────────────────────────────────────────────────────────────────────────
# NAME MATCHING (exact → strip middle name → fuzzy)
# ─────────────────────────────────────────────────────────────────────────────

def find_emp(name, date, emp_idx):
    n = norm(name)
    # 1. Exact
    if (n, date) in emp_idx:
        return emp_idx[(n, date)], 'exact', None
    # 2. Drop middle name (keep first + last word only)
    parts = n.split()
    if len(parts) >= 3:
        short = f"{parts[0]} {parts[-1]}"
        if (short, date) in emp_idx:
            return emp_idx[(short, date)], 'middle-name-stripped', f"'{name}' matched as '{emp_idx[(short,date)]['name']}'"
    # 3. Fuzzy match across all names on same date
    date_keys = [k for k in emp_idx if k[1] == date]
    if date_keys:
        emp_names = [k[0] for k in date_keys]
        close = difflib.get_close_matches(n, emp_names, n=1, cutoff=0.75)
        if close:
            match_key = (close[0], date)
            score = difflib.SequenceMatcher(None, n, close[0]).ratio()
            label = 'fuzzy-auto' if score >= 0.85 else 'fuzzy-review'
            return emp_idx[match_key], label, f"'{name}' fuzzy-matched to '{emp_idx[match_key]['name']}' ({score:.0%})"
    return None, 'no-match', None

# ─────────────────────────────────────────────────────────────────────────────
# RECONCILIATION ENGINE
# ─────────────────────────────────────────────────────────────────────────────

def reconcile(shifts, emp_idx):
    rows = []
    name_notes = []
    for s in shifts:
        date, emp, inv_hrs, worked_brk, lc = s['date'], s['emp'], s['inv_hrs'], s['worked_brk'], s['lc']
        if lc:
            rows.append({**s, 'e_in': None, 'e_out': None, 'e_raw': None, 'e_adj': None,
                          'start_diff': None, 'end_diff': None, 'hrs_diff': None,
                          'flag': 'LATE CANCEL', 'notes': '', 'match_type': 'lc'})
            continue
        m, match_type, match_note = find_emp(emp, date, emp_idx)
        notes = []
        if match_note:
            notes.append(match_note)
            name_notes.append(f"{date} {emp}: {match_note}")
        if worked_brk:
            notes.append("Worked break — full gross billed")
        if not m:
            rows.append({**s, 'e_in': None, 'e_out': None, 'e_raw': None, 'e_adj': None,
                          'start_diff': None, 'end_diff': None, 'hrs_diff': None,
                          'flag': 'NO PUNCH', 'notes': '; '.join(notes), 'match_type': 'no-match'})
            continue
        if m['split']:
            notes.append(f"Split punch ({len(m['entries'])} entries combined)")
        if m['raw'] < 1.0:
            notes.append(f"Micro-punch ({m['raw']}h — likely bad punch)")
        start_diff = round((m['in']  - s['start']).total_seconds() / 60, 1) if s['start'] else None
        end_diff   = round((m['out'] - s['end']  ).total_seconds() / 60, 1) if s['end']   else None
        hrs_diff   = round(inv_hrs - m['adj'], 2)
        if abs(hrs_diff) < 0.05:    flag = 'MATCH'
        elif hrs_diff > 0:          flag = f'OVERBILLED +{hrs_diff:.2f}h'
        else:                       flag = f'UNDERBILLED {hrs_diff:.2f}h'
        if worked_brk:  flag = 'WORKED BREAK'
        if m['raw'] < 1.0: flag = 'BAD PUNCH'
        if match_type == 'fuzzy-review': flag = f'REVIEW NAME MATCH | {flag}'
        rows.append({**s, 'e_in': m['in'], 'e_out': m['out'], 'e_raw': m['raw'],
                     'e_adj': m['adj'], 'start_diff': start_diff, 'end_diff': end_diff,
                     'hrs_diff': hrs_diff, 'flag': flag, 'notes': '; '.join(notes),
                     'match_type': match_type})
    return rows, name_notes

# ─────────────────────────────────────────────────────────────────────────────
# EXCEL BUILDERS
# ─────────────────────────────────────────────────────────────────────────────

def make_styles():
    return {
        'HDR_F': PatternFill("solid", start_color="1F4E79"),
        'HFONT': Font(name="Arial", bold=True, color="FFFFFF", size=9),
        'RED_F': PatternFill("solid", start_color="FFD7D7"),
        'YLW_F': PatternFill("solid", start_color="FFE699"),
        'ORG_F': PatternFill("solid", start_color="FCE4D6"),
        'GRN_F': PatternFill("solid", start_color="E2EFDA"),
        'GRY_F': PatternFill("solid", start_color="F2F2F2"),
        'UND_F': PatternFill("solid", start_color="FFF2CC"),
        'BLU_F': PatternFill("solid", start_color="DDEBF7"),
        'NF':  Font(name="Arial", size=9),
        'BF':  Font(name="Arial", bold=True, size=9),
        'RF':  Font(name="Arial", bold=True, size=9, color="C00000"),
        'BLF': Font(name="Arial", bold=True, size=9, color="1F4E79"),
        'CTR': Alignment(horizontal="center", vertical="center"),
        'LFT': Alignment(horizontal="left",   vertical="center"),
        'WRAP': Alignment(horizontal="left",  vertical="center", wrap_text=True),
        'TFmt': "M/D H:MM AM/PM",
        'RFILLS': {
            'RN':  PatternFill("solid", start_color="E2EFDA"),
            'LPN': PatternFill("solid", start_color="DDEBF7"),
            'LVN': PatternFill("solid", start_color="DDEBF7"),
            'CNA': PatternFill("solid", start_color="FCE4D6"),
            'CMA': PatternFill("solid", start_color="FFF2CC"),
        },
    }

def build_hours_excel(meta, shifts):
    S = make_styles()
    ROLE_ORDER = {'RN': 0, 'LPN': 1, 'LVN': 1, 'CMA': 2, 'CNA': 3}
    RATES = {'RN': 58, 'LPN': 45, 'LVN': 45, 'CMA': 28.5, 'CNA': 27}

    # Collect dates and employees
    all_dates = sorted(set(s['date'] for s in shifts))
    emp_role = {}
    for s in shifts:
        if s['emp'] not in emp_role:
            emp_role[s['emp']] = s['role']
    employees = sorted(emp_role, key=lambda n: (ROLE_ORDER.get(emp_role[n], 9), n))
    grid = defaultdict(lambda: defaultdict(list))
    for s in shifts:
        grid[s['emp']][s['date']].append(s['inv_hrs'])

    wb = Workbook(); ws = wb.active; ws.title = "Hours by Day"
    ncols = 2 + len(all_dates) + 2
    ws.merge_cells(f"A1:{get_column_letter(ncols)}1")
    ws["A1"] = f"Clipboard Invoice {meta['invoice']} — {meta['facility']} — {meta['period_start']}–{meta['period_end']}"
    ws["A1"].font = Font(name="Arial", bold=True, size=12, color="1F4E79")
    ws.row_dimensions[1].height = 22

    from calendar import day_abbr
    def day_label(d):
        try: dt = datetime.strptime(d + "/2026", "%m/%d/%Y"); return f"{dt.strftime('%a')} {d}"
        except: return d

    hdrs = ["Employee", "Role"] + [day_label(d) for d in all_dates] + ["Total Hrs", "Total $"]
    for ci, h in enumerate(hdrs, 1):
        c = ws.cell(row=2, column=ci, value=h)
        c.font = S['HFONT']; c.fill = S['HDR_F']; c.alignment = S['CTR']
    ws.row_dimensions[2].height = 18

    row = 3
    for emp in employees:
        role = emp_role[emp]
        rf = S['RFILLS'].get(role, PatternFill())
        ws.cell(row=row, column=1, value=emp).font = S['NF']
        ws.cell(row=row, column=1).alignment = S['LFT']
        ws.cell(row=row, column=1).fill = rf
        ws.cell(row=row, column=2, value=role).alignment = S['CTR']
        ws.cell(row=row, column=2).fill = rf
        ws.cell(row=row, column=2).font = Font(name="Arial", size=9, bold=(role == "RN"))
        th = 0; ta = 0
        for di, d in enumerate(all_dates):
            hrs_list = grid[emp].get(d, [])
            if hrs_list:
                hrs = sum(hrs_list); th += hrs; ta += hrs * RATES.get(role, 27)
                c = ws.cell(row=row, column=3+di, value=round(hrs, 2))
                c.fill = rf; c.alignment = S['CTR']; c.number_format = "0.00"; c.font = S['NF']
            else:
                ws.cell(row=row, column=3+di, value="").fill = rf
        bl = PatternFill("solid", start_color="D6E4F0")
        tc = ws.cell(row=row, column=3+len(all_dates), value=round(th, 2))
        tc.font = S['BF']; tc.alignment = S['CTR']; tc.fill = bl; tc.number_format = "0.00"
        tm = ws.cell(row=row, column=4+len(all_dates), value=round(ta, 2))
        tm.font = S['BF']; tm.alignment = Alignment(horizontal="right", vertical="center")
        tm.fill = bl; tm.number_format = '"$"#,##0.00'
        ws.row_dimensions[row].height = 15; row += 1

    tr = row
    ws.cell(row=tr, column=1, value="TOTAL").font = Font(name="Arial", bold=True, color="FFFFFF", size=9)
    for ci in range(1, ncols+1):
        ws.cell(row=tr, column=ci).fill = S['HDR_F']
    for ci in range(3, ncols+1):
        col = get_column_letter(ci)
        c = ws.cell(row=tr, column=ci, value=f"=SUM({col}3:{col}{tr-1})")
        c.font = Font(name="Arial", bold=True, color="FFFFFF", size=9)
        c.fill = S['HDR_F']; c.alignment = S['CTR']
        c.number_format = '"$"#,##0.00' if ci == ncols else "0.00"
    ws.row_dimensions[tr].height = 18

    ws.column_dimensions["A"].width = 28
    ws.column_dimensions["B"].width = 6
    for ci in range(3, 3+len(all_dates)): ws.column_dimensions[get_column_letter(ci)].width = 10
    ws.column_dimensions[get_column_letter(3+len(all_dates))].width = 11
    ws.column_dimensions[get_column_letter(4+len(all_dates))].width = 13

    buf = io.BytesIO(); wb.save(buf); buf.seek(0)
    return buf.read()

def build_flags_excel(meta, recon_rows):
    S = make_styles()

    def row_fill(flag):
        if 'NO PUNCH' in flag:    return S['RED_F']
        if 'BAD PUNCH' in flag:   return S['YLW_F']
        if 'WORKED BREAK' in flag: return S['YLW_F']
        if 'OVERBILLED' in flag:  return S['ORG_F']
        if 'UNDERBILLED' in flag: return S['UND_F']
        if 'REVIEW' in flag:      return S['BLU_F']
        if flag == 'MATCH':       return S['GRN_F']
        if flag == 'LATE CANCEL': return S['GRY_F']
        return S['GRN_F']

    wb = Workbook()

    # ── Sheet 1: Flags ────────────────────────────────────────────────────
    ws1 = wb.active; ws1.title = "Flags"
    ws1.merge_cells("A1:I1")
    ws1["A1"] = f"Invoice {meta['invoice']} — {meta['facility']} — Issues to Verify in Empion"
    ws1["A1"].font = Font(name="Arial", bold=True, size=12, color="1F4E79")
    ws1.row_dimensions[1].height = 22

    for ci, h in enumerate(["Date","Employee","Role","Issue","Invoice In","Invoice Out","Empion In","Empion Out","Notes"], 1):
        c = ws1.cell(row=2, column=ci, value=h)
        c.font = S['HFONT']; c.fill = S['HDR_F']; c.alignment = S['CTR']
    ws1.row_dimensions[2].height = 16

    ri = 3
    for r in recon_rows:
        if r['flag'] in ('MATCH', 'LATE CANCEL'): continue
        fill = row_fill(r['flag'])
        def s1(col, val, fmt=None, fnt=None, aln=None):
            c = ws1.cell(row=ri, column=col, value=val)
            c.fill = fill; c.alignment = aln or S['CTR']; c.font = fnt or S['NF']
            if val and fmt: c.number_format = fmt
        s1(1, r['date'], fnt=S['BF'])
        s1(2, r['emp'], aln=S['LFT'], fnt=S['BF'])
        s1(3, r['role'])
        fc = ws1.cell(row=ri, column=4, value=r['flag'])
        fc.fill = fill; fc.alignment = S['CTR']
        fc.font = S['RF'] if ('OVER' in r['flag'] or 'NO PUNCH' in r['flag']) else S['BF']
        s1(5, r['start'], S['TFmt']); s1(6, r['end'], S['TFmt'])
        s1(7, r['e_in'], S['TFmt']); s1(8, r['e_out'], S['TFmt'])
        for col in [5, 6, 7, 8]:
            c = ws1.cell(row=ri, column=col)
            if c.value is None: c.value = "—"
        s1(9, r['notes'], aln=S['WRAP'])
        ws1.row_dimensions[ri].height = 18; ri += 1

    ws1.column_dimensions["A"].width = 7;  ws1.column_dimensions["B"].width = 26
    ws1.column_dimensions["C"].width = 5;  ws1.column_dimensions["D"].width = 24
    for col in ["E","F","G","H"]: ws1.column_dimensions[col].width = 18
    ws1.column_dimensions["I"].width = 48
    ws1.freeze_panes = "A3"

    # ── Sheet 2: Punch-for-Punch ──────────────────────────────────────────
    ws2 = wb.create_sheet("Punch-for-Punch")
    ws2.merge_cells("A1:P1")
    ws2["A1"] = f"Invoice {meta['invoice']} — {meta['facility']} — Full Punch-for-Punch vs Empion"
    ws2["A1"].font = Font(name="Arial", bold=True, size=11, color="1F4E79")
    ws2.row_dimensions[1].height = 20

    for ci, h in enumerate(["Date","Employee","Role","Shift","Inv Start","Inv End","Inv Gross","Inv Billed",
                             "Emp In","Emp Out","Emp Raw","Emp Adj","Start Δ(min)","End Δ(min)","Hrs Diff","Flag / Notes"], 1):
        c = ws2.cell(row=2, column=ci, value=h)
        c.font = S['HFONT']; c.fill = S['HDR_F']
        c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    ws2.row_dimensions[2].height = 30

    for ri2, r in enumerate(recon_rows, 3):
        fill = row_fill(r['flag'])
        i_gross = round((r['end'] - r['start']).total_seconds() / 3600, 2) if r['start'] and r['end'] else None
        def cv2(col, val, fmt=None, fnt=None, aln=None):
            c = ws2.cell(row=ri2, column=col, value=val)
            c.fill = fill; c.alignment = aln or S['CTR']; c.font = fnt or S['NF']
            if fmt and val is not None: c.number_format = fmt
        note_val = r['flag'] + (' | ' + r['notes'] if r['notes'] else '')
        cv2(1, r['date']); cv2(2, r['emp'], aln=S['LFT']); cv2(3, r['role']); cv2(4, r['shift'])
        cv2(5, r['start'], S['TFmt']); cv2(6, r['end'], S['TFmt'])
        cv2(7, i_gross, "0.00"); cv2(8, r['inv_hrs'], "0.00", fnt=S['BF'])
        cv2(9, r['e_in'], S['TFmt']); cv2(10, r['e_out'], S['TFmt'])
        cv2(11, r['e_raw'], "0.00"); cv2(12, r['e_adj'], "0.00", fnt=S['BF'])
        sd = ws2.cell(row=ri2, column=13, value=r['start_diff'])
        sd.fill = fill; sd.alignment = S['CTR']; sd.number_format = "+0.0;-0.0;0"
        sd.font = S['RF'] if (r['start_diff'] and abs(r['start_diff']) >= 10) else S['NF']
        ed = ws2.cell(row=ri2, column=14, value=r['end_diff'])
        ed.fill = fill; ed.alignment = S['CTR']; ed.number_format = "+0.0;-0.0;0"
        ed.font = S['RF'] if (r['end_diff'] and abs(r['end_diff']) >= 10) else S['NF']
        hc = ws2.cell(row=ri2, column=15, value=r['hrs_diff'])
        hc.fill = fill; hc.alignment = S['CTR']; hc.number_format = "+0.00;-0.00;0.00"
        if r['hrs_diff'] and r.get('e_raw') and r['e_raw'] >= 1 and not r['worked_brk']:
            hc.font = S['RF'] if r['hrs_diff'] > 0.05 else (S['BLF'] if r['hrs_diff'] < -0.05 else S['NF'])
        else:
            hc.font = S['NF']
        fc = ws2.cell(row=ri2, column=16, value=note_val)
        fc.fill = fill; fc.alignment = S['WRAP']; fc.font = S['NF']
        ws2.row_dimensions[ri2].height = 14

    for i, w in enumerate([7,26,5,6,17,17,9,9,17,17,9,9,12,12,10,42], 1):
        ws2.column_dimensions[get_column_letter(i)].width = w
    ws2.freeze_panes = "A3"

    buf = io.BytesIO(); wb.save(buf); buf.seek(0)
    return buf.read()

# ─────────────────────────────────────────────────────────────────────────────
# STREAMLIT UI
# ─────────────────────────────────────────────────────────────────────────────

col1, col2 = st.columns(2)
with col1:
    pdf_file  = st.file_uploader("📄 Clipboard Invoice (PDF)", type="pdf")
with col2:
    xlsx_file = st.file_uploader("📊 Empion Punch Report (Excel)", type=["xlsx", "xls"])

if pdf_file and xlsx_file:
    if st.button("▶  Run Reconciliation", type="primary", use_container_width=True):
        with st.spinner("Parsing invoice..."):
            meta, shifts = parse_clipboard_invoice(pdf_file.read())
        with st.spinner("Parsing Empion punches..."):
            emp_idx = parse_empion(xlsx_file.read())
        with st.spinner("Reconciling..."):
            recon_rows, name_notes = reconcile(shifts, emp_idx)

        # ── Summary stats ─────────────────────────────────────────────────
        st.divider()
        st.subheader(f"Invoice {meta['invoice']} — {meta['facility']}")
        st.caption(f"Period: {meta['period_start']} – {meta['period_end']}  |  Balance Due: ${meta['balance_due']:,.2f}")

        total_items = len(recon_rows)
        n_match  = sum(1 for r in recon_rows if r['flag'] == 'MATCH')
        n_lc     = sum(1 for r in recon_rows if r['flag'] == 'LATE CANCEL')
        n_np     = sum(1 for r in recon_rows if 'NO PUNCH' in r['flag'])
        n_over   = sum(1 for r in recon_rows if 'OVERBILLED' in r['flag'])
        n_under  = sum(1 for r in recon_rows if 'UNDERBILLED' in r['flag'])
        n_other  = total_items - n_match - n_lc - n_np - n_over - n_under

        c1, c2, c3, c4, c5, c6 = st.columns(6)
        c1.metric("✅ Match",       n_match)
        c2.metric("🔴 No Punch",    n_np)
        c3.metric("🟠 Overbilled",  n_over)
        c4.metric("🟡 Underbilled", n_under)
        c5.metric("⚠️ Other",       n_other)
        c6.metric("⬜ Late Cancel", n_lc)

        # ── Name matching notes ───────────────────────────────────────────
        if name_notes:
            with st.expander(f"⚠️ {len(name_notes)} name(s) matched with fuzzy/middle-name logic — review"):
                for note in name_notes:
                    st.write(f"• {note}")

        # ── Flags table ───────────────────────────────────────────────────
        flag_rows = [r for r in recon_rows if r['flag'] not in ('MATCH', 'LATE CANCEL')]
        if flag_rows:
            st.subheader(f"🚩 {len(flag_rows)} Items to Review")
            import pandas as pd
            def fmt_dt(v):
                return v.strftime("%-m/%-d %-I:%M %p") if v else "—"
            table_data = [{
                "Date":       r['date'],
                "Employee":   r['emp'],
                "Role":       r['role'],
                "Issue":      r['flag'],
                "Inv In":     fmt_dt(r['start']),
                "Inv Out":    fmt_dt(r['end']),
                "Emp In":     fmt_dt(r['e_in']),
                "Emp Out":    fmt_dt(r['e_out']),
                "Notes":      r['notes'],
            } for r in flag_rows]
            st.dataframe(pd.DataFrame(table_data), use_container_width=True, hide_index=True)
        else:
            st.success("🎉 All shifts matched — no issues found!")

        # ── Downloads ─────────────────────────────────────────────────────
        st.divider()
        st.subheader("Download Reports")
        inv_slug = meta['invoice'].replace('-', '_')
        fac_slug = re.sub(r'[^a-zA-Z0-9]', '_', meta['facility'])[:20].strip('_')

        d1, d2 = st.columns(2)
        with d1:
            hours_bytes = build_hours_excel(meta, shifts)
            st.download_button(
                "⬇️  Hours by Employee (Excel)",
                data=hours_bytes,
                file_name=f"{fac_slug}_{inv_slug}_Hours.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True,
            )
        with d2:
            flags_bytes = build_flags_excel(meta, recon_rows)
            st.download_button(
                "⬇️  Flags & Punch-for-Punch (Excel)",
                data=flags_bytes,
                file_name=f"{fac_slug}_{inv_slug}_Flags.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True,
            )
else:
    st.info("Upload both files above to get started.")
