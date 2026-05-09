import pdfplumber
import pandas as pd
import re
from pathlib import Path
import tkinter as tk
from tkinter import filedialog, messagebox
import sys
import datetime
import win32clipboard
import html
from tkinter import ttk

GOOGLE_SHEET_ID = "1sPcHdIzRtsCnrknTyLKQ8QGM1gSw1B1GEa8BECSg4ds"

def clean_date(d):
    try:
        if not d: return ""
        d = d.replace(" ", "").strip()
        if re.match(r'\d{2}-[A-Za-z]{3}-\d{4}', d):
            return d
        if '/' in d:
            parts = d.split('/')
            if len(parts) >= 3:
                date_obj = datetime.datetime(int(parts[2]), int(parts[1]), int(parts[0]))
                return date_obj.strftime("%d-%b-%Y")
    except (ValueError, IndexError, TypeError):
        pass
    return d

def format_date_boe(d):
    try:
        if not d: return ""
        d = d.replace(" ", "").strip()
        if '/' in d:
            parts = d.split('/')
            if len(parts) >= 3:
                date_obj = datetime.datetime(int(parts[2]), int(parts[1]), int(parts[0]))
                return date_obj.strftime("%d-%b-%Y")
    except (ValueError, IndexError, TypeError):
        pass
    return d

def get_value_from_headers(words: list[dict], header_text: str, num_values: int, value_index_from_right: int) -> str:
    y_tolerance = 3
    lines = _group_words_into_lines(words, y_tolerance)
    sorted_tops = sorted(lines.keys())
    for i, top in enumerate(sorted_tops):
        line_words = sorted(lines[top], key=lambda w: w['x0'])
        line_text = ' '.join(w['text'] for w in line_words)
        if header_text in line_text:
            for j in range(i + 1, len(sorted_tops)):
                candidate = sorted(lines[sorted_tops[j]], key=lambda w: w['x0'])
                numbers = [w['text'] for w in candidate if re.search(r'\d', w['text'])]
                if not numbers:
                    continue
                if len(numbers) >= value_index_from_right:
                    return numbers[-value_index_from_right]
                break
    return ""

def clean_point_zero(val) -> str:
    if val is None or pd.isnull(val):
        return ""
    s = str(val).strip()
    if s.endswith(".0"):
        return s[:-2]
    return s

def _group_words_into_lines(words: list[dict], y_tolerance: int = 3) -> dict:
    lines: dict[int, list[dict]] = {}
    for w in words:
        top = round(w['top'] / y_tolerance) * y_tolerance
        if top not in lines:
            lines[top] = []
        lines[top].append(w)
    return lines


def _extract_invoices_from_page(page) -> list[str]:
    text = page.extract_text(layout=True)
    if not text or "2.INVOICE NO" not in text:
        return []
    invoices: list[str] = []
    in_inv = False
    inv_col = -1
    for line in text.split('\n'):
        if '2.INVOICE NO' in line and '3.INV. AMT' in line:
            in_inv = True
            inv_col = line.find('2.INVOICE NO')
            continue
        if in_inv:
            if re.match(r'^\s*\d+\s+', line):
                inv = line[max(0, inv_col - 2):inv_col + 20].strip()
                inv = inv.split()[0] if inv else ""
                if inv:
                    invoices.append(inv)
            elif line.strip() == '' or '1.EVENT' in line or '1.BUYER' in line:
                in_inv = False
    return invoices


def _extract_licence_from_page(page) -> tuple[float, bool, bool]:
    l_text = page.extract_text(layout=True)
    if not l_text:
        return 0.0, False, False
    subtotal = 0.0
    section_found = False
    in_lic = False
    section_ended = False
    for line in l_text.split('\n'):
        if '11.DEBIT DUTY' in line or 'DEBIT DUTY' in line:
            in_lic = True
            section_found = True
            continue
        if in_lic:
            if re.match(r'^\s*\d+\s+', line):
                parts = line.strip().split()
                val = parts[-1] if parts else ""
                try:
                    subtotal += float(val.replace(',', ''))
                except ValueError:
                    pass
            else:
                terminators = ('G. RE-IMPORT', 'GLOSSARY', 'Page ', 'H. CERTIFICATE')
                if any(t in line for t in terminators):
                    section_ended = True
                    break
    return subtotal, section_found, section_ended


def process_pdf(file_path: str, sr_no: int) -> tuple[dict, list[str], str]:
    data = {
        "SR. NO.": sr_no, "CHA JOB NO.": "", "SUPPLIER": "",
        "HAWB NO": "", "HAWB DT.": "", "BOE NO.": "", "BOE DATE": "",
        "INVOICE NO.": "", "SCRIPTS NO.": "", "SCRIPTS UTILISED VALUE": 0.0,
        "ASS. VALUE": "", "DUTY AMT. AS PAR BOE/CHECK LIST": "",
        "PENALTY CHARGES": "", "INTEREST CHARGES": "",
        "TOTAL DUTY AMT (CASH)": "", "REMARK FOR PENALTY & INTEREST": "NA",
        "RMS/NONRMS": "", "ETA": ""
    }
    warnings: list[str] = []
    fname = Path(file_path).name

    with pdfplumber.open(file_path) as pdf:
        num_pages = len(pdf.pages)
        p0 = pdf.pages[0]
        words_p0 = p0.extract_words()
        text_p0 = p0.extract_text(layout=False) or ""
        y_tolerance = 3
        lines_p0 = _group_words_into_lines(words_p0, y_tolerance)
        sorted_tops = sorted(lines_p0.keys())

        branch = "CSN"
        m_br = re.search(r'IEC/Br\s+[A-Z0-9]+/(\d)', text_p0)
        if m_br:
            br_code = m_br.group(1)
            if br_code == '5': branch = "CSN"
            elif br_code == '0': branch = "Pune"
            elif br_code == '2': branch = "CLC"
            else: warnings.append(f"File {fname}: Unknown branch code '{br_code}'. Defaulting to CSN.")
        else:
            warnings.append(f"File {fname}: IEC/Br not found. Defaulting to CSN.")

        be_no_x0, be_date_x0 = None, None
        for i, top in enumerate(sorted_tops):
            lw = sorted(lines_p0[top], key=lambda w: w['x0'])
            lt = ' '.join(w['text'] for w in lw)
            if 'Port' in lt and 'BE' in lt and 'No' in lt and 'Date' in lt:
                for hw in lw:
                    if hw['text'] == 'No' and be_no_x0 is None: be_no_x0 = hw['x0']
                    if hw['text'] == 'Date' and be_no_x0 is not None: be_date_x0 = hw['x0']
                for j in range(i + 1, min(i + 4, len(sorted_tops))):
                    vw = sorted(lines_p0[sorted_tops[j]], key=lambda w: w['x0'])
                    nums = [w for w in vw if re.search(r'\d', w['text'])]
                    if not nums: continue
                    for w in vw:
                        if re.match(r'^\d{7}$', w['text']) and be_no_x0 and abs(w['x0'] - be_no_x0) < 30:
                            data["BOE NO."] = w['text']
                        if re.match(r'\d{2}/\d{2}/\d{4}', w['text']) and be_date_x0 and abs(w['x0'] - be_date_x0) < 30:
                            data["BOE DATE"] = format_date_boe(w['text'])
                    break
                break
        if not data["BOE NO."]: warnings.append(f"File {fname}: BE No. not found.")
        if not data["BOE DATE"]: warnings.append(f"File {fname}: BE Date not found.")

        importer_name = ""
        for i, top in enumerate(sorted_tops):
            lw = sorted(lines_p0[top], key=lambda w: w['x0'])
            lt = ' '.join(w['text'] for w in lw)
            if '1.IMPORTER' in lt and 'NAME' in lt and 'ADDRESS' in lt:
                for j in range(i + 1, min(i + 8, len(sorted_tops))):
                    nw = sorted(lines_p0[sorted_tops[j]], key=lambda w: w['x0'])
                    candidate = ' '.join(w['text'] for w in nw).strip()
                    if len(candidate) > 4:
                        importer_name = candidate
                        break
                break
        data["IMPORTER"] = importer_name

        fcl_lcl = ""
        for i, top in enumerate(sorted_tops):
            lw = sorted(lines_p0[top], key=lambda w: w['x0'])
            lt = ' '.join(w['text'] for w in lw)
            if '2.LCL/' in lt and '1.SNO' in lt:
                lcl_header_x0 = next((w['x0'] for w in lw if '2.LCL/' in w['text']), None)
                if lcl_header_x0 is None: break
                has_l, has_f = False, False
                for j in range(i + 1, min(i + 6, len(sorted_tops))):
                    vw = sorted(lines_p0[sorted_tops[j]], key=lambda w: w['x0'])
                    for w in vw:
                        if abs(w['x0'] - lcl_header_x0) < 25:
                            if w['text'] in ('L', 'LCL'): has_l = True
                            elif w['text'] in ('F', 'FCL'): has_f = True
                if has_l: fcl_lcl = 'L'
                elif has_f: fcl_lcl = 'F'
                break
        if not fcl_lcl: warnings.append(f"File {fname}: FCL/LCL value not detected. HAWB NO left blank.")

        assess_val = ""
        for i, top in enumerate(sorted_tops):
            lw = sorted(lines_p0[top], key=lambda w: w['x0'])
            lt = ' '.join(w['text'] for w in lw)
            if '8.ASSESS' in lt and '1.BE' in lt:
                assess_x0 = next((w['x0'] for w in lw if w['text'] == '8.ASSESS'), None)
                if assess_x0 is None: break
                for j in range(i + 1, min(i + 8, len(sorted_tops))):
                    vw = sorted(lines_p0[sorted_tops[j]], key=lambda w: w['x0'])
                    for w in vw:
                        if w['text'] in ('Y', 'N', 'P') and abs(w['x0'] - assess_x0) < 25:
                            assess_val = w['text']
                            break
                    if assess_val: break
                break
        if assess_val == 'N': data["RMS/NONRMS"] = "RMS"
        elif assess_val == 'Y': data["RMS/NONRMS"] = "NONRMS"
        else: warnings.append(f"File {fname}: 8.ASSESS (RMS/NONRMS) not detected.")

        data["ASS. VALUE"] = get_value_from_headers(words_p0, "18.TOT.ASS VAL", 8, 1)
        tot_amt = get_value_from_headers(words_p0, "14.TOTAL DUTY", 5, 1)
        if tot_amt: data["TOTAL DUTY AMT (CASH)"] = tot_amt
        pnlty = get_value_from_headers(words_p0, "14.TOTAL DUTY", 5, 3)
        if pnlty and pnlty != "0": data["PENALTY CHARGES"] = pnlty
        int_c = get_value_from_headers(words_p0, "14.TOTAL DUTY", 5, 4)
        if int_c and int_c != "0": data["INTEREST CHARGES"] = int_c
        tot_duty = get_value_from_headers(words_p0, "14.TOTAL DUTY", 5, 5)
        if tot_duty: data["DUTY AMT. AS PAR BOE/CHECK LIST"] = tot_duty

        awb_no, awb_dt = "", ""
        for i, top in enumerate(sorted_tops):
            line_words = sorted(lines_p0[top], key=lambda w: w['x0'])
            line_text = ' '.join(w['text'] for w in line_words)
            if "1.IGM" in line_text and "6.MAWB" in line_text:
                value_words = []
                for j in range(i + 1, min(i + 8, len(sorted_tops))):
                    candidate = sorted(lines_p0[sorted_tops[j]], key=lambda w: w['x0'])
                    cand_text = ' '.join(w['text'] for w in candidate)
                    if "1.BOND" in cand_text or "1.SR NO" in cand_text: break
                    if any(re.search(r'[A-Za-z0-9]', w['text']) for w in candidate):
                        value_words.extend(candidate)
                if not value_words: break
                def _get_col_value(header_prefix: str) -> str:
                    h_words, found = [], False
                    for hw in line_words:
                        if hw['text'].startswith(header_prefix.split('.')[0] + '.' + header_prefix.split('.')[1].split()[0] if '.' in header_prefix else header_prefix): found = True
                        if found:
                            h_words.append(hw)
                            if len(h_words) > 1 and re.match(r'\d+\.', hw['text']) and hw != h_words[0]:
                                h_words.pop(); break
                    if not h_words: return ""
                    hx0, hx1 = h_words[0]['x0'], h_words[-1]['x1']
                    vals = [nw['text'] for nw in value_words if nw['x0'] >= hx0 - 15 and nw['x1'] <= hx1 + 15]
                    return ''.join(vals)
                if fcl_lcl == 'F':
                    awb_no = _get_col_value("6.MAWB")
                    mawb_x = next((hw['x0'] for hw in line_words if '6.MAWB' in hw['text']), 300)
                    dts = [w['text'] for w in value_words if re.match(r'\d{2}/\d{2}/\d{4}', w['text']) and w['x0'] > mawb_x]
                    if dts: awb_dt = dts[0]
                elif fcl_lcl == 'L':
                    awb_no = _get_col_value("8.HAWB")
                    hawb_x = next((hw['x0'] for hw in line_words if '8.HAWB' in hw['text']), 400)
                    dts = [w['text'] for w in value_words if re.match(r'\d{2}/\d{2}/\d{4}', w['text']) and w['x0'] > hawb_x]
                    if dts: awb_dt = dts[0]
                break
        data["HAWB NO"] = awb_no
        data["HAWB DT."] = clean_date(awb_dt)
        if not awb_no: warnings.append(f"File {fname}: Correct HAWB/MAWB number not selected.")

        invoices: list[str] = []
        inv_header_x0, inv_header_x1, inv_amt_x0 = None, None, None
        for i_top, top in enumerate(sorted_tops):
            lw = sorted(lines_p0[top], key=lambda w: w['x0'])
            lt = ' '.join(w['text'] for w in lw)
            if '2.INVOICE' in lt and ('3.INV.' in lt or '3.INV' in lt):
                for hw in lw:
                    if '2.INVOICE' in hw['text']: inv_header_x0 = hw['x0']
                    if hw['text'] == 'NO' and inv_header_x0 is not None and hw['x0'] > inv_header_x0: inv_header_x1 = hw['x1']
                    if '3.INV.' in hw['text'] or hw['text'] == '3.INV': inv_amt_x0 = hw['x0']
                if inv_header_x0 is None: break
                if inv_header_x1 is None: inv_header_x1 = inv_header_x0 + 60
                if inv_amt_x0 is None: inv_amt_x0 = inv_header_x1 + 50
                for j in range(i_top + 1, len(sorted_tops)):
                    row_words = sorted(lines_p0[sorted_tops[j]], key=lambda w: w['x0'])
                    row_text = ' '.join(w['text'] for w in row_words)
                    if '1.EVENT' in row_text or 'Submission' in row_text: break
                    for rw in row_words:
                        if rw['x0'] >= inv_header_x0 - 15 and rw['x1'] <= inv_amt_x0 - 5:
                            if re.match(r'^\d+$', rw['text']): invoices.append(rw['text'])
                break
        if invoices: data["INVOICE NO."] = ", ".join(invoices)
        else: warnings.append(f"File {fname}: Invoice numbers not extracted.")

        supplier = ""
        if num_pages > 1:
            words_p1 = pdf.pages[1].extract_words()
            lines_p1 = _group_words_into_lines(words_p1, y_tolerance)
            tops_p1 = sorted(lines_p1.keys())
            for target_header in ("SUPPLIER NAME & ADDRESS", "THIRD PARTY NAME & ADDRESS", "SELLER'S NAME & ADDRESS"):
                found = False
                for i, top in enumerate(tops_p1):
                    lt = ' '.join(w['text'] for w in sorted(lines_p1[top], key=lambda w: w['x0']))
                    if target_header in lt:
                        if i + 1 < len(tops_p1):
                            sup_words = sorted(lines_p1[tops_p1[i + 1]], key=lambda w: w['x0'])
                            candidate = ' '.join(w['text'] for w in sup_words if w['x0'] < 300)
                            if candidate: supplier = candidate; found = True
                        break
                if found: break
        if supplier: data["SUPPLIER"] = supplier
        else: warnings.append(f"File {fname}: Supplier name not found.")

        lic_tot, lic_found = 0.0, False
        last_lic_page = -1
        for p_idx in range(num_pages - 1, 1, -1):
            p_text = pdf.pages[p_idx].extract_text()
            if p_text and ("11.DEBIT DUTY" in p_text or "DEBIT DUTY" in p_text or "LICENCE DETAILS" in p_text):
                last_lic_page = p_idx; break
        lic_start_page = last_lic_page
        if last_lic_page >= 0:
            for p_idx in range(last_lic_page - 1, 1, -1):
                p_text = pdf.pages[p_idx].extract_text()
                if p_text and ("11.DEBIT DUTY" in p_text or "DEBIT DUTY" in p_text or "LICENCE DETAILS" in p_text):
                    lic_start_page = p_idx
                else: break
        if lic_start_page >= 0:
            for p_idx in range(lic_start_page, last_lic_page + 1):
                subtotal, found_on_page, ended = _extract_licence_from_page(pdf.pages[p_idx])
                lic_tot += subtotal
                if found_on_page: lic_found = True
        if lic_found: data["SCRIPTS UTILISED VALUE"] = round(lic_tot, 2)
        else: data["SCRIPTS UTILISED VALUE"] = 0; warnings.append(f"File {fname}: Licence Details section not found. SCRIPTS UTILISED VALUE set to 0.")

        if not data["ASS. VALUE"]: warnings.append(f"File {fname}: Total assessable value not extracted.")
        if not data["DUTY AMT. AS PAR BOE/CHECK LIST"]: warnings.append(f"File {fname}: Total duty not extracted.")
        if not data["TOTAL DUTY AMT (CASH)"]: warnings.append(f"File {fname}: Total cash duty not extracted.")

    return data, warnings, branch


# ═══════════════════════════════════════════════════════════════
# CHECKLIST PARSER
# ═══════════════════════════════════════════════════════════════

_CL_BRANCH_MAP = {"5": "CSN", "0": "Pune", "2": "CLC"}

_SKODA_KEYWORDS = ["SKODA", "VOLKSWAGEN"]

def _is_skoda(importer_name: str) -> bool:
    upper = importer_name.upper()
    return any(kw in upper for kw in _SKODA_KEYWORDS)

def _to_float(v) -> float:
    try: return float(str(v).replace(',', '').strip())
    except (ValueError, TypeError): return 0.0

def normalize_bl(val: str) -> str:
    if not val: return ""
    return re.sub(r'[^A-Z0-9]', '', str(val).upper())


def parse_checklist(file_path: str) -> tuple[dict, list[str]]:
    warnings: list[str] = []
    fname = Path(file_path).name
    data: dict = {
        "job_no": "", "be_no": "", "branch": "", "bl_no": "",
        "assessable_value": 0.0, "duty_amount": 0.0, "foregone_duty": 0.0,
        "file_name": fname,
    }

    with pdfplumber.open(file_path) as pdf:
        if not pdf.pages:
            warnings.append(f"Checklist {fname}: PDF has no pages.")
            return data, warnings
        num_pages = len(pdf.pages)
        first_text = pdf.pages[0].extract_text(layout=True) or ""

        m = re.search(r"Job\s*No\s+IR/(\d+)/", first_text)
        if m: data["job_no"] = m.group(1)
        else: warnings.append(f"Checklist {fname}: Job No not found.")

        m = re.search(r"BE\s*No\.?\s*/\s*Date\s+(\d{7})\b", first_text, re.IGNORECASE)
        if not m:
            m = re.search(r"BE\s*No\.?\s*.*?\b(\d{7})\b", first_text, re.IGNORECASE)
            if m:
                start, end = m.start(1), m.end(1)
                if (start > 0 and first_text[start - 1].isdigit()) or (end < len(first_text) and first_text[end].isdigit()):
                    m = None
        if m: data["be_no"] = m.group(1)
        else: warnings.append(f"Checklist {fname}: BE No not found. Will use BL/HBL fallback if available.")

        m = re.search(r"Branch\s*SNo:\s*(\d+)", first_text, re.IGNORECASE)
        if m: data["branch"] = _CL_BRANCH_MAP.get(m.group(1), f"Unknown({m.group(1)})")
        else: warnings.append(f"Checklist {fname}: Branch SNo not found.")

        m_hbl = re.search(r"HBL\s*No\.?\s+([A-Z0-9/-]+)\s+dt\.", first_text, re.IGNORECASE)
        m_bl = re.search(r"(?<!H)BL\s*No\.?\s+([A-Z0-9/-]+)\s+dt\.", first_text, re.IGNORECASE)
        data["bl_no"] = (m_hbl.group(1) if m_hbl else "") or (m_bl.group(1) if m_bl else "")
        if not data["bl_no"]: warnings.append(f"Checklist {fname}: BL/HBL No not found.")

        gross_found, foregone_found = False, False
        for page_idx, page in enumerate(reversed(pdf.pages)):
            actual_page_num = num_pages - page_idx
            if gross_found and foregone_found: break

            # --- PERFORMANCE FIX: Fast Pre-Scan ---
            fast_text = page.extract_text(layout=False) or ""
            
            if not foregone_found and "BCD Foregone Duty" in fast_text:
                text = page.extract_text(layout=True) or ""
                matches = list(re.finditer(r"BCD\s+Foregone\s+Duty\s+([\d,]+\.?\d*)", text))
                if matches:
                    try:
                        data["foregone_duty"] = float(matches[-1].group(1).replace(',', ''))
                        foregone_found = True
                    except ValueError: pass

            if not gross_found and "BE" in fast_text and "Gross" in fast_text and "Total" in fast_text:
                page_words = page.extract_words()
                lines = _group_words_into_lines(page_words, y_tolerance=3)
                for top in sorted(lines.keys(), reverse=True):
                    lw = sorted(lines[top], key=lambda w: w['x0'])
                    texts = [w['text'] for w in lw]
                    joined = ' '.join(texts)
                    if 'BE' not in texts or 'Gross' not in texts or 'Total' not in texts: continue
                    if 'Page' in texts and 'BE' not in joined.split('Page')[0]: continue
                    gross_x = next((w['x0'] for w in lw if w['text'] == 'Gross'), None)
                    if gross_x is None: continue
                    left_nums, right_nums = [], []
                    for w in lw:
                        try:
                            v = float(w['text'].replace(',', ''))
                            (left_nums if w['x0'] < gross_x else right_nums).append(v)
                        except ValueError: pass
                    if left_nums and right_nums:
                        data["assessable_value"] = left_nums[-1]
                        data["duty_amount"] = right_nums[-1]
                        gross_found = True
                        break

        if not gross_found: warnings.append(f"Checklist {fname}: BE Gross Total not found.")
        if not foregone_found: warnings.append(f"Checklist {fname}: BCD Foregone Duty not found.")

    return data, warnings

def _detect_doc_type(file_path: str) -> str:
    """
    Detects whether a PDF is a Checklist, BE Copy, or unknown.
    Uses only the first page. Returns 'checklist', 'be', or 'unknown'.
    """
    try:
        with pdfplumber.open(file_path) as pdf:
            if not pdf.pages:
                return "unknown"

            first_page = pdf.pages[0]
            text = first_page.extract_text(layout=False) or ""
            words = first_page.extract_words()

            # ── Checklist signals ──
            cl_signals = 0

            # Signal 1: "CheckList" appears in the page title area
            # (top 15% of page height, i.e. small y values)
            page_height = first_page.height
            title_words = [w["text"] for w in words if w["top"] < page_height * 0.15]
            if any("CheckList" in w or "CHECKLIST" in w for w in title_words):
                cl_signals += 1

            # Signal 2: Job No in IR/NNNNN/NN-NN format
            if re.search(r"Job\s*No\s+IR/\d+/", text):
                cl_signals += 1

            # Signal 3: Branch SNo: followed by a digit
            if re.search(r"Branch\s*SNo\s*:\s*\d", text):
                cl_signals += 1

            # ── BE Copy signals ──
            be_signals = 0

            # Signal 1: IEC/Br code line
            if re.search(r"IEC/Br\s+[A-Z0-9]+/\d", text):
                be_signals += 1

            # Signal 2: "Port Code BE No BE Date" as column headers on one line
            # Check using word positions — these three appear left-to-right on same line
            lines = {}
            for w in words:
                top = round(w["top"] / 3) * 3
                lines.setdefault(top, []).append(w["text"])
            header_found = any(
                "Port" in line_words and "BE" in line_words and "No" in line_words and "Date" in line_words
                for line_words in lines.values()
            )
            if header_found:
                be_signals += 1

            # Signal 3: Numbered field headers unique to BE structure
            if "14.TOTAL DUTY" in text or "18.TOT.ASS VAL" in text:
                be_signals += 1

            # ── Decision ──
            # Require at least 2 signals to be confident
            if cl_signals >= 2 and cl_signals > be_signals:
                return "checklist"
            if be_signals >= 2 and be_signals > cl_signals:
                return "be"
            return "unknown"

    except Exception:
        return "unknown"

def validate_be_vs_checklist(be_data: dict, cl_data: dict, be_branch: str, check_branch: bool = True) -> dict:
    result: dict = {
        "be_no": be_data.get("BOE NO.", ""), "job_no": cl_data.get("job_no", ""),
        "file_name": cl_data.get("file_name", ""), "status": "OK", "match_method": "BE No",
        "branch_match": True, "bl_match": True,
        "be_ass": 0.0, "cl_ass": 0.0, "ass_diff": 0.0,
        "be_duty": 0.0, "cl_duty": 0.0, "duty_diff": 0.0,
        "be_scripts": 0.0, "cl_foregone": 0.0, "foregone_diff": 0.0,
        "remarks": [],
    }
    if check_branch and cl_data["branch"] and cl_data["branch"] != be_branch:
        result["branch_match"] = False
        result["remarks"].append(f"Branch mismatch: BE={be_branch}, CL={cl_data['branch']}")
    be_bl = str(be_data.get("HAWB NO", "")).strip()
    cl_bl = str(cl_data.get("bl_no", "")).strip()
    if cl_bl and be_bl and normalize_bl(cl_bl) != normalize_bl(be_bl):
        result["bl_match"] = False
        result["remarks"].append(f"BL mismatch: BE={be_bl}, CL={cl_bl}")

    result["be_ass"] = _to_float(be_data.get("ASS. VALUE", 0))
    result["cl_ass"] = cl_data.get("assessable_value", 0.0)
    result["ass_diff"] = round(result["be_ass"] - result["cl_ass"], 2)
    result["be_duty"] = _to_float(be_data.get("TOTAL DUTY AMT (CASH)", 0))
    result["cl_duty"] = cl_data.get("duty_amount", 0.0)
    result["duty_diff"] = round(result["be_duty"] - result["cl_duty"], 2)
    result["be_scripts"] = _to_float(be_data.get("SCRIPTS UTILISED VALUE", 0))
    result["cl_foregone"] = cl_data.get("foregone_duty", 0.0)
    result["foregone_diff"] = round(result["be_scripts"] - result["cl_foregone"], 2)

    has_high = (abs(result["ass_diff"]) > 500 or abs(result["duty_diff"]) > 500 or abs(result["foregone_diff"]) > 500)
    has_warn  = (100 < abs(result["ass_diff"]) <= 500 or
                100 < abs(result["duty_diff"]) <= 500 or
                100 < abs(result["foregone_diff"]) <= 500)
    has_minor = (0 < abs(result["ass_diff"]) <= 100 or
                0 < abs(result["duty_diff"]) <= 100 or
                0 < abs(result["foregone_diff"]) <= 100)
    has_verify = not result["branch_match"] or not result["bl_match"]

    if has_high: result["status"] = "HIGH RISK"
    elif has_verify or has_warn: result["status"] = "WARNING"
    elif has_minor: result["status"] = "MINOR"
    else: result["status"] = "OK"
    return result


def get_unified_be_data(row: dict, branch: str) -> dict:
    if branch == "CSN":
        return {"BOE NO.": str(row.get("BOE NO.", "")).strip(), "HAWB NO": str(row.get("HAWB NO", "")).strip(),
                "ASS. VALUE": str(row.get("ASS. VALUE", "0")).strip(), "TOTAL DUTY AMT (CASH)": str(row.get("TOTAL DUTY AMT (CASH)", "0")).strip(),
                "SCRIPTS UTILISED VALUE": str(row.get("SCRIPTS UTILISED VALUE", "0")).strip()}
    elif branch == "CLC":
        return {"BOE NO.": str(row.get("BOE NO.", "")).strip(), "HAWB NO": str(row.get("MBL NO/HBL NO", "")).strip(),
                "ASS. VALUE": str(row.get("ASS. VALUE", "0")).strip(), "TOTAL DUTY AMT (CASH)": str(row.get("TOTAL DUTY AMT (CASH)", "0")).strip(),
                "SCRIPTS UTILISED VALUE": str(row.get("SCRIPTS UTILISED VALUE", "0")).strip()}
    elif branch == "Pune":
        return {"BOE NO.": str(row.get("BE No", "")).strip(), "HAWB NO": str(row.get("BL No", "")).strip(),
                "ASS. VALUE": str(row.get("Ass. Value", "0")).strip(), "TOTAL DUTY AMT (CASH)": str(row.get("Total Amt", "0")).strip(),
                "SCRIPTS UTILISED VALUE": str(row.get("Debit Duty", "0")).strip()}
    elif branch == "General":
        return {"BOE NO.": str(row.get("BOE NO.", "")).strip(), "HAWB NO": str(row.get("HAWB NO", "")).strip(),
                "ASS. VALUE": str(row.get("ASS. VALUE", "0")).strip(), "TOTAL DUTY AMT (CASH)": str(row.get("TOTAL DUTY AMT (CASH)", "0")).strip(),
                "SCRIPTS UTILISED VALUE": str(row.get("SCRIPTS UTILISED VALUE", "0")).strip()}
    return {}


def format_csn(data): return {"SR. NO.": data["SR. NO."], "CHA JOB NO.": data["CHA JOB NO."], "SUPPLIER": data["SUPPLIER"], "HAWB NO": data["HAWB NO"], "HAWB DT.": data["HAWB DT."], "BOE NO.": data["BOE NO."], "BOE DATE": data["BOE DATE"], "INVOICE NO.": data["INVOICE NO."], "SCRIPTS NO.": data["SCRIPTS NO."], "SCRIPTS UTILISED VALUE": data["SCRIPTS UTILISED VALUE"], "ASS. VALUE": data["ASS. VALUE"], "DUTY AMT. AS PAR BOE/CHECK LIST": data["DUTY AMT. AS PAR BOE/CHECK LIST"], "PENALTY CHARGES": data["PENALTY CHARGES"], "INTEREST CHARGES": data["INTEREST CHARGES"], "TOTAL DUTY AMT (CASH)": data["TOTAL DUTY AMT (CASH)"], "REMARK FOR PENALTY & INTEREST": data["REMARK FOR PENALTY & INTEREST"], "RMS/NONRMS": data["RMS/NONRMS"]}
def format_clc(data): return {"SR. NO.": data["SR. NO."], "CHA JOB NO.": data["CHA JOB NO."], "SUPPLIER": data["SUPPLIER"], "MBL NO/HBL NO": data["HAWB NO"], "HAWB DT.": data["HAWB DT."], "BOE NO.": data["BOE NO."], "BOE DATE": data["BOE DATE"], "INVOICE NO.": data["INVOICE NO."], "SCRIPTS NO.": data["SCRIPTS NO."], "SCRIPTS UTILISED VALUE": data["SCRIPTS UTILISED VALUE"], "ASS. VALUE": data["ASS. VALUE"], "DUTY AMT. AS PAR BOE/CHECK LIST": data["DUTY AMT. AS PAR BOE/CHECK LIST"], "PENALTY CHARGES": data["PENALTY CHARGES"], "INTEREST CHARGES": data["INTEREST CHARGES"], "TOTAL DUTY AMT (CASH)": data["TOTAL DUTY AMT (CASH)"], "REMARK FOR PENALTY & INTEREST": data["REMARK FOR PENALTY & INTEREST"]}
def format_pune(data): return {"Sr.No.": data["SR. NO."], "Job No": data["CHA JOB NO."], "Supplier": data["SUPPLIER"], "Invoice no.": data["INVOICE NO."], "BL No": data["HAWB NO"], "BL DATE": data["HAWB DT."], "BE No": data["BOE NO."], "BE Dt.": data["BOE DATE"], "Ass. Value": data["ASS. VALUE"], "Debit Duty": data["SCRIPTS UTILISED VALUE"], "Duty": data["DUTY AMT. AS PAR BOE/CHECK LIST"], "Penalty": data["PENALTY CHARGES"], "Interest Charges": data["INTEREST CHARGES"], "Total Amt": data["TOTAL DUTY AMT (CASH)"], "REMARK": data["REMARK FOR PENALTY & INTEREST"], "ETA": data["ETA"]}

def format_general(data: dict) -> dict:
    return {
        "SR. NO.":                          data.get("SR. NO.", ""),
        "IMPORTER":                         data.get("IMPORTER", ""),
        "SUPPLIER":                         data.get("SUPPLIER", ""),
        "HAWB NO":                          data.get("HAWB NO", ""),
        "HAWB DT.":                         data.get("HAWB DT.", ""),
        "BOE NO.":                          data.get("BOE NO.", ""),
        "BOE DATE":                         data.get("BOE DATE", ""),
        "INVOICE NO.":                      data.get("INVOICE NO.", ""),
        "SCRIPTS UTILISED VALUE":           data.get("SCRIPTS UTILISED VALUE", ""),
        "ASS. VALUE":                       data.get("ASS. VALUE", ""),
        "DUTY AMT. AS PAR BOE/CHECK LIST":  data.get("DUTY AMT. AS PAR BOE/CHECK LIST", ""),
        "TOTAL DUTY AMT (CASH)":            data.get("TOTAL DUTY AMT (CASH)", ""),
        "REMARK FOR PENALTY & INTEREST":    data.get("REMARK FOR PENALTY & INTEREST", ""),
    }


_NUMERIC_COLS = {"SR. NO.", "SCRIPTS UTILISED VALUE", "ASS. VALUE", "DUTY AMT. AS PAR BOE/CHECK LIST", "PENALTY CHARGES", "INTEREST CHARGES", "TOTAL DUTY AMT (CASH)", "Sr.No.", "Debit Duty", "Ass. Value", "Duty", "Penalty", "Total Amt", "Interest Charges"}


def generate_html_table(df: pd.DataFrame, exclude_cols: set[str] | None = None) -> str:
    cols = [c for c in df.columns if not exclude_cols or c not in exclude_cols]
    html_out = '<table border="1" cellspacing="0" cellpadding="4" style="border-collapse:collapse; border:1px solid black; font-family:Calibri,Arial,sans-serif; font-size:11pt;">\n  <thead>\n    <tr style="background-color:#ADD8E6;">\n'
    for col in cols:
        html_out += f'      <th style="border:1px solid black; text-align:center; font-weight:bold; vertical-align:middle;">{col}</th>\n'
    html_out += '    </tr>\n  </thead>\n  <tbody>\n'
    for _, row in df.iterrows():
        html_out += '    <tr>\n'
        for col in cols:
            val = html.escape(str(row[col])) if pd.notnull(row[col]) else ""
            align = "right" if col in _NUMERIC_COLS else "left"
            html_out += f'      <td style="border:1px solid black; text-align:{align}; vertical-align:middle;">{val}</td>\n'
        html_out += '    </tr>\n'
    html_out += '  </tbody>\n</table>'
    return html_out


def copy_html_to_clipboard(html_str: str):
    header = "Version:0.9\r\nStartHTML:{0:08d}\r\nEndHTML:{1:08d}\r\nStartFragment:{2:08d}\r\nEndFragment:{3:08d}\r\n"
    prefix = "<html><body><!--StartFragment-->"
    suffix = "<!--EndFragment--></body></html>"
    html_bytes = html_str.encode('utf-8')
    prefix_bytes = prefix.encode('utf-8')
    suffix_bytes = suffix.encode('utf-8')
    start_html = 105
    start_fragment = start_html + len(prefix_bytes)
    end_fragment = start_fragment + len(html_bytes)
    end_html = end_fragment + len(suffix_bytes)
    header_str = header.format(start_html, end_html, start_fragment, end_fragment)
    clipboard_data = header_str.encode('utf-8') + prefix_bytes + html_bytes + suffix_bytes
    win32clipboard.OpenClipboard()
    try:
        win32clipboard.EmptyClipboard()
        cf_html = win32clipboard.RegisterClipboardFormat("HTML Format")
        win32clipboard.SetClipboardData(cf_html, clipboard_data)
    finally:
        win32clipboard.CloseClipboard()


# ═══════════════════════════════════════════════════════════════
# GUI
# ═══════════════════════════════════════════════════════════════

_PRIMARY_BLUE = "#1F3F6E"
_ACCENT_RED = "#D8232A"
_DARK_TEXT = "#1E1E1E"
_MUTED_GRAY = "#6B7280"
_LIGHT_BG = "#F4F6F8"
_PANEL_WHITE = "#FFFFFF"
_BORDER_GRAY = "#E5E7EB"
_HOVER_BLUE = "#2A528F"
_HEADER_BG = "#D6E4F0"

def get_base_path() -> Path:
    return Path(sys._MEIPASS) if getattr(sys, 'frozen', False) else Path(__file__).parent

def get_exe_path() -> Path:
    return Path(sys.executable).parent if getattr(sys, 'frozen', False) else Path(__file__).parent

LOGO_PATH = get_base_path() / "logo.png"


class BECopyParserApp:

    def __init__(self, root: tk.Tk) -> None:
        self.root = root
        self.root.title("BE Copy Parser — Nagarkot Forwarders Pvt. Ltd.")
        self.root.configure(bg=_LIGHT_BG)
        self.root.state("zoomed")
        self.root.minsize(1024, 600)

        self.be_warnings: list[str] = []
        self.cl_warnings: list[str] = []
        self.all_warnings: list[str] = []
        self.raw_be_data: dict[str, dict] = {}
        self.be_branches: dict[str, str] = {}
        self.checklist_data: dict[str, dict] = {}
        self.checklist_data_by_bl: dict[str, dict] = {}
        self.validation_results: dict[str, dict] = {}
        self.branch_dfs: dict[str, pd.DataFrame] = {}
        self._copy_flash_id = None

        self._build_header()
        self._build_body()
        self._build_footer()

    def _build_header(self) -> None:
        header = tk.Frame(self.root, bg=_PANEL_WHITE, bd=0, highlightthickness=1, highlightbackground=_BORDER_GRAY)
        header.pack(fill=tk.X, side=tk.TOP)
        header.pack_propagate(False)
        header.configure(height=85)
        try:
            self._logo_img = tk.PhotoImage(file=str(LOGO_PATH))
            factor = max(1, self._logo_img.height() // 30)
            self._logo_img = self._logo_img.subsample(factor)
            tk.Label(header, image=self._logo_img, bg=_PANEL_WHITE).pack(side=tk.LEFT, padx=20)
        except Exception: pass
        title_frame = tk.Frame(header, bg=_PANEL_WHITE)
        title_frame.place(relx=0.5, rely=0.5, anchor=tk.CENTER)
        tk.Label(title_frame, text="Duty Call Automation", font=("Segoe UI", 18, "bold"), fg=_PRIMARY_BLUE, bg=_PANEL_WHITE).pack()
        tk.Label(title_frame, text="Automated Bill of Entry & Checklist Reconciliation", font=("Segoe UI", 10), fg=_MUTED_GRAY, bg=_PANEL_WHITE).pack()

    def _build_body(self) -> None:
        body = tk.Frame(self.root, bg=_LIGHT_BG)
        body.pack(fill=tk.BOTH, expand=True, padx=16, pady=(12, 4))

        # ── Row 1: Toolbar with buttons + status message on the right ──
        toolbar = tk.Frame(body, bg=_LIGHT_BG)
        toolbar.pack(fill=tk.X, pady=(0, 6))

        self.btn_select = self._brand_button(toolbar, "📂  Select BE PDFs", self._on_select_files)
        self.btn_select.pack(side=tk.LEFT)
        self.btn_checklist = self._brand_button(toolbar, "📄  Select Checklists", self._on_select_checklists, state=tk.DISABLED)
        self.btn_checklist.pack(side=tk.LEFT, padx=(8, 0))
        self.btn_clear = self._brand_button(toolbar, "Clear All", self._on_clear, state=tk.DISABLED)
        self.btn_clear.pack(side=tk.LEFT, padx=(8, 0))
        self.btn_copy_table = self._brand_button(toolbar, "📋  Copy Table", self._on_copy_table, state=tk.DISABLED)
        self.btn_copy_table.pack(side=tk.RIGHT)

        # Status label in toolbar row, right-aligned, high contrast
        self.status_var = tk.StringVar(value="Select BE Copy PDFs to begin")
        self.status_lbl = tk.Label(
            toolbar, textvariable=self.status_var,
            font=("Segoe UI", 10, "bold"), fg=_PRIMARY_BLUE, bg=_LIGHT_BG,
            anchor=tk.E, padx=12,
        )
        self.status_lbl.pack(side=tk.RIGHT, padx=(0, 8))

        # ── Row 2: Validation summary bar (hidden until validation runs) ──
        self._summary_frame = tk.Frame(body, bg=_LIGHT_BG)
        self._summary_frame.pack(fill=tk.X, pady=(0, 4))
        self.summary_bar = tk.Label(
            self._summary_frame, text="", font=("Segoe UI", 10, "bold"),
            padx=12, pady=6, anchor=tk.W,
        )
        # Not packed yet — shown after validation

        # ── Row 3: Table area ──
        table_frame = tk.Frame(body, bg=_PANEL_WHITE, bd=1, relief=tk.SOLID)
        table_frame.pack(fill=tk.BOTH, expand=True)
        self.tree = None
        self._table_frame = table_frame

        # REMOVED: Warnings panel (warn_frame / warn_text) — not needed

    def _build_footer(self) -> None:
        footer = tk.Frame(self.root, bg=_PANEL_WHITE, height=28, highlightthickness=1, highlightbackground=_BORDER_GRAY)
        footer.pack(fill=tk.X, side=tk.BOTTOM)
        footer.pack_propagate(False)
        tk.Label(footer, text="Nagarkot Forwarders Pvt. Ltd. ©", font=("Segoe UI", 8), fg=_MUTED_GRAY, bg=_PANEL_WHITE).pack(side=tk.LEFT, padx=12)

    def _brand_button(self, parent, text, command, state=tk.NORMAL):
        btn = tk.Button(parent, text=text, command=command, state=state,
                        font=("Segoe UI", 10, "bold"), fg="#FFFFFF", bg=_PRIMARY_BLUE,
                        activebackground=_HOVER_BLUE, activeforeground="#FFFFFF",
                        bd=0, padx=14, pady=6, cursor="hand2", relief=tk.FLAT)
        def _on_enter(e):
            if btn['state'] != tk.DISABLED: btn.configure(bg=_HOVER_BLUE)
        def _on_leave(e):
            if btn['state'] != tk.DISABLED: btn.configure(bg=_PRIMARY_BLUE)
        btn.bind("<Enter>", _on_enter)
        btn.bind("<Leave>", _on_leave)
        return btn

    # ── Status helpers ─────────────────────────────────────────
    def _set_status(self, text: str, color: str = _PRIMARY_BLUE) -> None:
        """Updates status label with high-contrast text."""
        self.status_var.set(text)
        self.status_lbl.configure(fg=color)
        self.root.update_idletasks()

    def _flash_copy_success(self) -> None:
        """Shows a prominent temporary green flash near the Copy button."""
        if self._copy_flash_id:
            self.root.after_cancel(self._copy_flash_id)
        self._set_status("✅  Copied to clipboard! Paste into Outlook.", "#2E7D32")
        # Revert after 5 seconds
        self._copy_flash_id = self.root.after(5000, lambda: self._set_status("Ready", _MUTED_GRAY))

    def _update_summary_bar(self) -> None:
        if not self.validation_results:
            self.summary_bar.pack_forget()
            return
        ok = sum(1 for v in self.validation_results.values() if v.get("status") == "OK")
        minor = sum(1 for v in self.validation_results.values() if v.get("status") == "MINOR")
        warn = sum(1 for v in self.validation_results.values() if v.get("status") == "WARNING")
        high = sum(1 for v in self.validation_results.values() if v.get("status") == "HIGH RISK")
        total = len(self.validation_results)
        if high > 0:
            txt = f"⚠  {high} HIGH RISK, {warn} Warning, {ok} OK — out of {total} BE(s). Review required!"
            bg, fg = "#FFEBEE", "#C62828"
        elif warn > 0 or minor > 0:
            txt = f"⚠  {warn} Warning, {minor} Minor, {ok} OK — out of {total} BE(s)."
            bg, fg = "#FFF8E1", "#F57F17"
        else:
            txt = f"✅  All {total} BE(s) verified OK."
            bg, fg = "#E8F5E9", "#2E7D32"
        self.summary_bar.configure(text=txt, bg=bg, fg=fg)
        self.summary_bar.pack(fill=tk.X)

    # ── File selection ─────────────────────────────────────────
    def _on_select_files(self) -> None:
        paths = filedialog.askopenfilenames(title="Select BE Copy PDFs", filetypes=[("PDF files", "*.pdf")])
        if not paths:
            return

        # Validate file types before processing
        wrong, unknown = [], []
        for fp in paths:
            doc_type = _detect_doc_type(fp)
            if doc_type == "checklist":
                wrong.append(Path(fp).name)
            elif doc_type == "unknown":
                unknown.append(Path(fp).name)

        if wrong:
            messagebox.showerror(
                "Wrong File Type",
                "These appear to be Checklist PDFs, not BE Copies:\n\n"
                + "\n".join(wrong)
                + "\n\nPlease use 'Select Checklists' for these."
            )
            return

        if unknown:
            proceed = messagebox.askyesno(
                "Unrecognised File(s)",
                "Could not confirm the type of these files:\n\n"
                + "\n".join(unknown)
                + "\n\nThey may be valid BE Copies with unusual formatting.\n"
                "Proceed anyway?"
            )
            if not proceed:
                return

        self._process(list(paths))

    # ── CHANGE 1: Merge checklists instead of reset ──────────
    def _on_select_checklists(self) -> None:
        paths = filedialog.askopenfilenames(title="Select Checklist PDFs", filetypes=[("PDF files", "*.pdf")])
        if not paths: return

        # Validate file types before doing any heavy parsing
        wrong, unknown = [], []
        for fp in paths:
            doc_type = _detect_doc_type(fp)
            if doc_type == "be":
                wrong.append(Path(fp).name)
            elif doc_type == "unknown":
                unknown.append(Path(fp).name)

        if wrong:
            messagebox.showerror(
                "Wrong File Type",
                f"These appear to be BE Copy PDFs, not Checklists:\n\n"
                + "\n".join(wrong)
                + "\n\nPlease use 'Select BE PDFs' for these."
            )
            return  # Hard stop — don't process anything from this batch

        if unknown:
            proceed = messagebox.askyesno(
                "Unrecognised File(s)",
                f"Could not confirm the type of these files:\n\n"
                + "\n".join(unknown)
                + "\n\nThey may be valid checklists with unusual formatting.\n"
                "Proceed anyway?"
            )
            if not proceed:
                return

        self._set_status(f"Parsing {len(paths)} checklist(s)…")

        # NOTE: We do NOT reset self.checklist_data or self.checklist_data_by_bl here.
        # New checklists are merged in: new keys are added, duplicate keys are replaced.
        self.cl_warnings = []
        new_count = 0
        replaced_count = 0

        for fp in paths:
            self._set_status(f"Parsing: {Path(fp).name}…")
            try:
                cl_data, cl_warnings = parse_checklist(str(fp))
                self.cl_warnings.extend(cl_warnings)
                be_no = str(cl_data.get("be_no", "")).strip()
                bl_key = normalize_bl(cl_data.get("bl_no", ""))
                stored = False

                # Hard block: BE No is mandatory — no fallback allowed
                if not be_no:
                    messagebox.showwarning(
                        "BE No. Missing",
                        f"The following checklist does not have a BE No.:\n\n"
                        f"{Path(fp).name}\n\n"
                    )
                    # Skip this file entirely — do not store, do not affect other checklists
                    continue

                # BE No present — store normally
                if be_no in self.checklist_data:
                    self.cl_warnings.append(f"Checklist {Path(fp).name}: Replacing existing checklist for BE No {be_no}.")
                    replaced_count += 1
                else:
                    new_count += 1
                self.checklist_data[be_no] = cl_data

                # Also index by BL for cross-reference lookups (not for matching — BE No is always primary)
                if bl_key:
                    self.checklist_data_by_bl[bl_key] = cl_data


            except Exception as e:
                self.cl_warnings.append(f"Checklist {Path(fp).name}: ERROR — {e}")

        self._run_validation()
        self._render_table()
        self._update_summary_bar()

        if self.branch_dfs:
            self.btn_copy_table.configure(state=tk.NORMAL)
            # Build a helpful status message
            total_loaded = len(self.checklist_data) + len([
                k for k in self.checklist_data_by_bl
                if not any(normalize_bl(cl.get("bl_no", "")) == k
                           for cl in self.checklist_data.values())
            ])
            parts = []
            if new_count: parts.append(f"{new_count} added")
            if replaced_count: parts.append(f"{replaced_count} replaced")
            action_text = ", ".join(parts) if parts else "processed"
            self._set_status(f"Checklists {action_text}. Total loaded: {total_loaded}. Review, then copy.")
            if hasattr(self, "_tab_keys") and "__validation_report__" in self._tab_keys:
                try: self.notebook.select(self._tab_keys.index("__validation_report__"))
                except Exception: pass
        
        # import threading
        # threading.Thread(target=self._upload_high_diff_to_sheets, daemon=True).start()

    def _on_clear(self) -> None:
        self.branch_dfs = {}
        self.raw_be_data = {}
        self.be_branches = {}
        self.checklist_data = {}
        self.checklist_data_by_bl = {}
        self.validation_results = {}
        self.be_warnings, self.cl_warnings, self.all_warnings = [], [], []
        self.summary_bar.pack_forget()
        self.btn_checklist.configure(state=tk.DISABLED)
        self.btn_copy_table.configure(state=tk.DISABLED)
        self.btn_clear.configure(state=tk.DISABLED)
        for child in self._table_frame.winfo_children():
            child.destroy()
        self.notebook = None
        self._set_status("Cleared")

    def _process(self, file_paths: list) -> None:
        self._set_status(f"Processing {len(file_paths)} PDF(s)…")
        branch_dfs_data = {
            "CSN": self.branch_dfs.get("CSN", pd.DataFrame()).to_dict("records") if "CSN" in self.branch_dfs else [],
            "CLC": self.branch_dfs.get("CLC", pd.DataFrame()).to_dict("records") if "CLC" in self.branch_dfs else [],
            "Pune": self.branch_dfs.get("Pune", pd.DataFrame()).to_dict("records") if "Pune" in self.branch_dfs else [],
            "General": self.branch_dfs.get("General", pd.DataFrame()).to_dict("records") if "General" in self.branch_dfs else [],
        }
        self.be_warnings, self.cl_warnings, self.all_warnings = [], [], []
        self.validation_results = {}
        total_parsed = 0
        skoda_count = 0
        gen_count = 0
        for i, fp in enumerate(file_paths):
            self._set_status(f"Processing {Path(fp).name}  ({i+1}/{len(file_paths)})…")
            try:
                data, warnings, branch = process_pdf(str(fp), i + 1)
                
                importer = data.get("IMPORTER", "")
                if not importer:
                    warnings.append(f"File {Path(fp).name}: Importer name not found. Defaulting to General profile.")
                
                data["_party"] = "Skoda" if _is_skoda(importer) else "General"
                
                if data["_party"] == "Skoda":
                    effective_branch = branch if branch in branch_dfs_data and branch != "General" else "CSN"
                    skoda_count += 1
                else:
                    effective_branch = "General"
                    gen_count += 1
                    
                be_no = str(data.get("BOE NO.", "")).strip()
                if be_no: self.raw_be_data[be_no] = data; self.be_branches[be_no] = effective_branch
                
                if effective_branch == "CSN": fd = format_csn(data)
                elif effective_branch == "CLC": fd = format_clc(data)
                elif effective_branch == "Pune": fd = format_pune(data)
                else: fd = format_general(data)
                
                fd["SR. NO." if effective_branch != "Pune" else "Sr.No."] = len(branch_dfs_data[effective_branch]) + 1
                branch_dfs_data[effective_branch].append(fd)
                self.be_warnings.extend(warnings); total_parsed += 1
            except Exception as e:
                self.be_warnings.append(f"File {Path(fp).name}: ERROR — {e}")
        self.branch_dfs = {}
        for b, d in branch_dfs_data.items():
            if d:
                df = pd.DataFrame(d)
                self.branch_dfs[b] = df.map(clean_point_zero) if hasattr(df, 'map') else df.applymap(clean_point_zero)
        self.all_warnings = self.be_warnings + self.cl_warnings
        self._render_table()
        self._update_summary_bar()
        if self.branch_dfs:
            self.btn_checklist.configure(state=tk.NORMAL)
            self.btn_clear.configure(state=tk.NORMAL)
            msg = f"Parsed {total_parsed} BE(s) ({skoda_count} Skoda, {gen_count} General). Select checklists for validation report."
            self._set_status(msg)
        else:
            self._set_status("No data parsed.", _ACCENT_RED)

    def _run_validation(self) -> None:
        self.validation_results = {}
        self.cl_warnings = [w for w in self.cl_warnings if "No matching BE" not in w]
        for branch, df in self.branch_dfs.items():
            be_col = "BOE NO." if branch in ("CSN", "CLC", "General") else "BE No" if branch == "Pune" else None
            job_col = "CHA JOB NO." if branch in ("CSN", "CLC", "General") else "Job No" if branch == "Pune" else None
            for idx, row in df.iterrows():
                be_data = get_unified_be_data(row.to_dict(), branch)
                be_no = be_data["BOE NO."]
                if not be_no: continue
                cl = self.checklist_data.get(be_no)
                match_method = "BE No"
                if cl:
                    check_branch = (branch != "General")
                    result = validate_be_vs_checklist(be_data, cl, branch, check_branch=check_branch)
                    result["match_method"] = match_method
                    if match_method == "BL/HBL fallback" and result["status"] == "OK":
                        result["status"] = "WARNING"
                    self.validation_results[be_no] = result
                    if cl.get("job_no") and job_col: df.at[idx, job_col] = cl["job_no"]
                else:
                    self.validation_results[be_no] = {
                        "be_no": be_no, "job_no": "", "file_name": "", "status": "HIGH RISK",
                        "match_method": "No match", "remarks": ["No matching checklist found"],
                        "be_ass": _to_float(be_data["ASS. VALUE"]), "cl_ass": 0.0, "ass_diff": 0.0,
                        "be_duty": _to_float(be_data["TOTAL DUTY AMT (CASH)"]), "cl_duty": 0.0, "duty_diff": 0.0,
                        "be_scripts": _to_float(be_data["SCRIPTS UTILISED VALUE"]), "cl_foregone": 0.0, "foregone_diff": 0.0,
                        "branch_match": True, "bl_match": True}
        for be_no, cl in self.checklist_data.items():
            if be_no not in self.validation_results:
                self.cl_warnings.append(f"Checklist BE {be_no} (Job {cl.get('job_no','')}): No matching BE Copy.")
        self.all_warnings = self.be_warnings + self.cl_warnings

    # ── TABLE ──────────────────────────────────────────────────
    def _render_table(self) -> None:
        selected_idx = 0
        if hasattr(self, "notebook") and self.notebook:
            try: selected_idx = self.notebook.index(self.notebook.select())
            except Exception: pass
        for child in self._table_frame.winfo_children(): child.destroy()
        if not self.branch_dfs:
            self.notebook = None; return

        style = ttk.Style(); style.theme_use("clam")
        style.configure("BE.Treeview", background=_PANEL_WHITE, foreground=_DARK_TEXT, fieldbackground=_PANEL_WHITE, font=("Segoe UI", 10), rowheight=28)
        style.configure("BE.Treeview.Heading", background=_HEADER_BG, foreground=_PRIMARY_BLUE, font=("Segoe UI", 9, "bold"))
        style.map("BE.Treeview", background=[("selected", _HOVER_BLUE)], foreground=[("selected", "#FFFFFF")])

        self.notebook = ttk.Notebook(self._table_frame)
        self.notebook.pack(fill=tk.BOTH, expand=True)
        self.branch_trees = {}
        self._tab_keys: list[str] = []

        for branch, df in self.branch_dfs.items():
            if df.empty: continue
            if self.validation_results:
                be_col = "BOE NO." if "BOE NO." in df.columns else "BE No" if "BE No" in df.columns else None
                statuses, levels = [], []
                for _, row in df.iterrows():
                    be_no = str(row.get(be_col, "")).strip() if be_col else ""
                    vr = self.validation_results.get(be_no, {})
                    st = vr.get("status", "")
                    if not st:
                        statuses.append("")
                        levels.append("")
                        continue

                    # Build user-friendly remark for the Validation column
                    remarks_list = []
                    if vr.get("match_method") == "BL/HBL fallback":
                        remarks_list.append("Matched via BL/HBL")
                    elif vr.get("match_method") == "No match":
                        remarks_list.append("No matching checklist found")
                    ass_diff = vr.get("ass_diff", 0.0)
                    duty_diff = vr.get("duty_diff", 0.0)
                    foregone_diff = vr.get("foregone_diff", 0.0)
                    if ass_diff != 0.0:
                        remarks_list.append(f"Assessable Value Diff: {ass_diff}")
                    if duty_diff != 0.0:
                        remarks_list.append(f"Duty Diff: {duty_diff}")
                    if foregone_diff != 0.0:
                        remarks_list.append(f"Foregone Duty Diff: {foregone_diff}")
                    for r in vr.get("remarks", []):
                        if "mismatch" in r.lower():
                            remarks_list.append(r)

                    if st == "OK" and not remarks_list:
                        remark_str = "Verified OK - All values match"
                    else:
                        remark_str = f"{st}"
                        if remarks_list:
                            remark_str += f" — {', '.join(remarks_list)}"
                    statuses.append(remark_str)
                    levels.append(st)
                df = df.copy(); df["Validation"] = statuses; df["_lvl"] = levels
                self.branch_dfs[branch] = df

            frame = tk.Frame(self.notebook, bg=_PANEL_WHITE)
            self.notebook.add(frame, text=f"  {branch} ({len(df)})  ")
            self._tab_keys.append(branch)
            display_cols = [c for c in df.columns if c != "_lvl"]
            if "Validation" in display_cols: display_cols.remove("Validation"); display_cols.insert(0, "Validation")

            xscroll = ttk.Scrollbar(frame, orient=tk.HORIZONTAL)
            yscroll = ttk.Scrollbar(frame, orient=tk.VERTICAL)
            tree = ttk.Treeview(frame, columns=display_cols, show="headings", style="BE.Treeview", xscrollcommand=xscroll.set, yscrollcommand=yscroll.set)
            xscroll.configure(command=tree.xview); yscroll.configure(command=tree.yview)
            yscroll.pack(side=tk.RIGHT, fill=tk.Y); xscroll.pack(side=tk.BOTTOM, fill=tk.X)
            tree.pack(fill=tk.BOTH, expand=True)
            self.branch_trees[branch] = tree
            tree.bind("<Double-1>", lambda e, t=tree, b=branch: self._on_double_click_cell(e, t, b))

            for col in display_cols:
                a = tk.E if col in _NUMERIC_COLS else tk.W
                w = 350 if col == "Validation" else 60 if ("NO" in col.upper() and "SR" in col.upper()) else 130
                if col == "Validation": a = tk.W
                tree.heading(col, text=col, anchor=tk.CENTER); tree.column(col, width=w, minwidth=60, anchor=a)

            for idx, (_, row) in enumerate(df.iterrows()):
                vals = [str(row[c]) if pd.notnull(row[c]) else "" for c in display_cols]
                lvl = str(row.get("_lvl", ""))
                tag = ("high_risk" if lvl == "HIGH RISK" else
                        "warning"   if lvl == "WARNING"   else
                        "minor"     if lvl == "MINOR"     else
                        "ok"        if lvl == "OK"        else
                        ("even" if idx % 2 == 0 else "odd"))
                tree.insert("", tk.END, values=vals, tags=(tag,))
            tree.tag_configure("even", background=_PANEL_WHITE); tree.tag_configure("odd", background="#EEF2F7")
            tree.tag_configure("ok", background="#E8F5E9"); tree.tag_configure("warning", background="#FFF8E1")
            tree.tag_configure("high_risk", background="#FFEBEE")
            tree.tag_configure("minor", background="#FFF3E0")

        if self.validation_results: self._render_validation_report_tab(style)
        if hasattr(self, "notebook") and self.notebook:
            try: self.notebook.select(selected_idx)
            except Exception: pass

    # ── CHANGE 3: Validation Report — removed Match Method & Status columns ──
    def _render_validation_report_tab(self, style) -> None:
        frame = tk.Frame(self.notebook, bg=_PANEL_WHITE)
        self.notebook.add(frame, text="  Validation Report  ")
        self._tab_keys.append("__validation_report__")
        report_cols = ["BE No.", "CHA Job No.", "Checklist File",
                       "BE Ass. Value", "CL Ass. Value", "Ass. Value Diff",
                       "BE Duty (Cash)", "CL Duty (Cash)", "Duty Diff",
                       "BE Scripts Value", "CL Foregone Duty", "Foregone Duty Diff", "Remarks"]
        canvas = tk.Canvas(frame, bg=_PANEL_WHITE, highlightthickness=0)
        yscroll = ttk.Scrollbar(frame, orient=tk.VERTICAL, command=canvas.yview)
        xscroll = ttk.Scrollbar(frame, orient=tk.HORIZONTAL, command=canvas.xview)
        sc = tk.Frame(canvas, bg=_PANEL_WHITE)
        sc.bind("<Configure>", lambda e: canvas.configure(scrollregion=canvas.bbox("all")))
        canvas.create_window((0, 0), window=sc, anchor="nw")
        canvas.configure(yscrollcommand=yscroll.set, xscrollcommand=xscroll.set)
        yscroll.pack(side=tk.RIGHT, fill=tk.Y); xscroll.pack(side=tk.BOTTOM, fill=tk.X)
        canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        
        group_definitions = [
            ("doc",      "Document Details",          3, "#CFD8DC"),
            ("ass",      "Assessable Value Breakdown", 3, "#BBDEFB"),
            ("duty",     "Duty Amount (Cash)",         3, "#C8E6C9"),
            ("scripts",  "Scripts & Foregone Duty",    3, "#E1BEE7"),
            ("remarks",  "Status",                     1, "#CFD8DC"),
        ]

        column_mapping = {
            0:  ("doc", 0),
            1:  ("doc", 1),
            2:  ("doc", 2),
            3:  ("ass", 0),
            4:  ("ass", 1),
            5:  ("ass", 2),
            6:  ("duty", 0),
            7:  ("duty", 1),
            8:  ("duty", 2),
            9:  ("scripts", 0),
            10: ("scripts", 1),
            11: ("scripts", 2),
            12: ("remarks", 0),
        }

        group_frames = {}
        for col_idx, (name, title, span, bg_color) in enumerate(group_definitions):
            outer = tk.Frame(sc, bg=_PRIMARY_BLUE, padx=2, pady=0)  # Blue border
            outer.grid(row=0, column=col_idx, sticky="nsew")
            inner = tk.Frame(outer, bg=bg_color)
            inner.pack(fill=tk.BOTH, expand=True)
            group_frames[name] = inner

            # Draw the Super Header inside each group inner frame
            tk.Label(inner, text=title, font=("Segoe UI", 10, "bold"), fg=_DARK_TEXT, bg=bg_color,
                     padx=12, pady=6, borderwidth=1, highlightthickness=1, highlightbackground=_BORDER_GRAY
                     ).grid(row=0, column=0, columnspan=span, sticky="nsew")

        # Draw the regular sub-headers at row=1 inside each group inner frame
        for ci, col in enumerate(report_cols):
            group_name, inner_col = column_mapping[ci]
            tk.Label(group_frames[group_name], text=col, font=("Segoe UI", 9, "bold"), fg=_PRIMARY_BLUE, bg=_HEADER_BG,
                     padx=12, pady=8, borderwidth=1, highlightthickness=1, highlightbackground=_BORDER_GRAY
                     ).grid(row=1, column=inner_col, sticky="nsew")

        col_widths = [90, 80, 220, 120, 120, 100, 120, 120, 100, 120, 120, 100, 200]
        for ci, width in enumerate(col_widths):
            group_name, inner_col = column_mapping[ci]
            group_frames[group_name].grid_columnconfigure(inner_col, minsize=width)

        # Print data inside each group inner frame starting at row=2
        for ri, (be_no, vr) in enumerate(self.validation_results.items(), start=2):
            bg = _PANEL_WHITE if ri % 2 == 0 else "#EEF2F7"
            vals = [vr.get("be_no",""), vr.get("job_no",""), vr.get("file_name",""),
                    vr.get("be_ass",""), vr.get("cl_ass",""), vr.get("ass_diff",0.0),
                    vr.get("be_duty",""), vr.get("cl_duty",""), vr.get("duty_diff",0.0),
                    vr.get("be_scripts",""), vr.get("cl_foregone",""), vr.get("foregone_diff",0.0),
                    "; ".join(vr.get("remarks",[]))]
            for ci, val in enumerate(vals):
                group_name, inner_col = column_mapping[ci]
                cn = report_cols[ci]
                cbg, cfg, fw = bg, _DARK_TEXT, "normal"
                if cn in ("Ass. Value Diff", "Duty Diff", "Foregone Duty Diff"):
                    try: v = float(val)
                    except: v = 0.0
                    if v != 0:
                        if abs(v) <= 100:
                            cbg, cfg, fw = "#FFF3E0", "#E65100", "normal"
                        elif abs(v) <= 500:
                            cbg, cfg, fw = "#FFE082", "#E65100", "bold"
                        else:
                            cbg, cfg, fw = "#FFCDD2", "#B71C1C", "bold"
                tk.Label(group_frames[group_name], text=str(val), font=("Segoe UI", 9, fw), fg=cfg, bg=cbg,
                         padx=12, pady=6, borderwidth=1, highlightthickness=1, highlightbackground=_BORDER_GRAY
                         ).grid(row=ri, column=inner_col, sticky="nsew")

    def _upload_high_diff_to_sheets(self) -> None:
        """Silently appends HIGH RISK / No-match validation rows to Google Sheet."""
        try:
            import gspread
            import google.auth
        except ImportError:
            self.be_warnings.append("Google Sheets upload skipped: gspread or google-auth not installed.")
            return

        high_diff_rows = [
            vr for vr in self.validation_results.values()
            if abs(vr.get("ass_diff", 0)) > 500
            or abs(vr.get("duty_diff", 0)) > 500
            or abs(vr.get("foregone_diff", 0)) > 500
            or vr.get("match_method") == "No match"
        ]
        if not high_diff_rows:
            return  # Nothing to upload — exit silently

        try:
            import os
            from pathlib import Path
            exe_dir = get_exe_path()
            cred_file = exe_dir / "credentials.json"
            client_secret = exe_dir / "client_secret.json"
            token_file = exe_dir / "token.json"
            
            if cred_file.exists():
                gc = gspread.service_account(filename=str(cred_file))
            elif client_secret.exists() or token_file.exists():
                gc = gspread.oauth(
                    credentials_filename=str(client_secret),
                    authorized_user_filename=str(token_file)
                )
            else:
                scopes = ["https://www.googleapis.com/auth/spreadsheets"]
                creds, _ = google.auth.default(scopes=scopes)
                gc = gspread.authorize(creds)
                
            sh = gc.open_by_key(GOOGLE_SHEET_ID)
            ws = sh.worksheet("Validation Log")

            timestamp = datetime.datetime.now().strftime("%d-%b-%Y %H:%M")

            rows_to_append = []
            for vr in high_diff_rows:
                be_no = vr.get("be_no", "")
                importer = self.raw_be_data.get(be_no, {}).get("IMPORTER", "Unknown")
                rows_to_append.append([
                    timestamp,                              # Timestamp
                    importer,                               # Importer
                    vr.get("be_no", ""),                    # BE No
                    vr.get("job_no", ""),                   # Job No
                    vr.get("be_ass", ""),                   # BE Ass Value
                    vr.get("cl_ass", ""),                   # CL Ass Value
                    vr.get("ass_diff", ""),                 # Ass Value Diff
                    vr.get("be_duty", ""),                  # BE Duty
                    vr.get("cl_duty", ""),                  # CL Duty
                    vr.get("duty_diff", ""),                # Duty Diff
                    vr.get("be_scripts", ""),               # BE Scripts
                    vr.get("cl_foregone", ""),              # CL Foregone
                    vr.get("foregone_diff", ""),            # Foregone Diff
                    vr.get("status", ""),                   # Status
                    "; ".join(vr.get("remarks", [])),       # Remarks
                ])

            ws.append_rows(rows_to_append, value_input_option="USER_ENTERED")
            def _update_ui():
                self._set_status(
                    f"{self.status_var.get()} — {len(rows_to_append)} high-diff row(s) logged to sheet.",
                    "#2E7D32"
                )
            self.root.after(0, _update_ui)

        except Exception as e:
            err_msg = str(e)
            if "insufficient authentication scopes" in err_msg:
                err_msg = "Google Auth scopes missing. Run gcloud auth application-default login --scopes=https://www.googleapis.com/auth/spreadsheets"
            def _fail_ui():
                self._set_status(f"Google Sheets Upload Failed: {err_msg}", "#D8232A")
            self.root.after(0, _fail_ui)

    def _on_double_click_cell(self, event, tree, branch) -> None:
        row_id = tree.identify_row(event.y); col_id = tree.identify_column(event.x)
        if not row_id or not col_id: return
        ci = int(col_id[1:]) - 1; cols = list(tree["columns"]); cn = cols[ci]
        if cn.upper() in ("SR. NO.", "SR.NO.") or cn == "Validation": return
        bbox = tree.bbox(row_id, col_id)
        if not bbox: return
        x, y, w, h = bbox
        entry = tk.Entry(tree, font=("Segoe UI", 10), justify=tk.RIGHT if cn in _NUMERIC_COLS else tk.LEFT)
        entry.insert(0, tree.item(row_id, "values")[ci]); entry.select_range(0, tk.END); entry.focus_set()
        entry.place(x=x, y=y, width=w, height=h)
        def _save(e=None):
            if not entry.winfo_exists(): return
            nv = entry.get(); vals = list(tree.item(row_id, "values")); vals[ci] = nv; tree.item(row_id, values=vals)
            ri = tree.index(row_id); df = self.branch_dfs[branch]
            if cn in df.columns: df.iat[ri, df.columns.get_loc(cn)] = nv
            entry.destroy()
            if self.checklist_data or self.checklist_data_by_bl:
                self._run_validation(); self._render_table(); self._update_summary_bar()
        entry.bind("<Return>", _save); entry.bind("<FocusOut>", _save)

    def _get_active_branch(self):
        if not hasattr(self, "notebook") or not self.notebook or not hasattr(self, "_tab_keys"): return None
        idx = self.notebook.index(self.notebook.select())
        return self._tab_keys[idx] if idx < len(self._tab_keys) else None

    def _show_confirm(self, severity, details) -> bool:
        dlg = tk.Toplevel(self.root); dlg.withdraw(); dlg.title("Confirm Copy")
        dlg.configure(bg=_PANEL_WHITE); dlg.resizable(False, False)
        dlg.update_idletasks()
        
        x = self.root.winfo_x() + (self.root.winfo_width() - 620) // 2
        y = self.root.winfo_y() + (self.root.winfo_height() - 500) // 2
        dlg.geometry(f"620x500+{x}+{y}"); dlg.transient(self.root); dlg.grab_set()
        
        result = tk.BooleanVar(value=False)
        hbg = _ACCENT_RED if severity == "HIGH RISK" else "#F57F17"
        hf = tk.Frame(dlg, bg=hbg, pady=12); hf.pack(fill=tk.X)
        tk.Label(hf, text=f"{'🚨' if severity == 'HIGH RISK' else '⚠️'}  REVIEW REQUIRED", font=("Segoe UI", 12, "bold"), fg="#FFF", bg=hbg).pack()
        
        cf = tk.Frame(dlg, bg=_PANEL_WHITE, padx=20, pady=15); cf.pack(fill=tk.BOTH, expand=True)
        txt = tk.Text(cf, font=("Consolas", 10), height=12, bg="#F9FAFB", fg=_DARK_TEXT, relief=tk.SOLID, bd=1, padx=10, pady=10)
        txt.pack(fill=tk.BOTH, expand=True, pady=(0, 12))
        
        txt.tag_configure("hdr", font=("Consolas", 10, "bold"), foreground=_PRIMARY_BLUE)
        txt.tag_configure("val", font=("Consolas", 10, "bold"), foreground=_ACCENT_RED)
        
        for line in details.split("\n"):
            if line.startswith("BE "): 
                txt.insert(tk.END, line + "\n", "hdr")
            elif ":" in line and line.startswith("  - "): 
                p = line.split(":", 1); txt.insert(tk.END, p[0] + ":", ""); txt.insert(tk.END, p[1] + "\n", "val")
            else: 
                txt.insert(tk.END, line + "\n")
                
        txt.configure(state=tk.DISABLED)
        tk.Label(cf, text="Do you still want to copy?", font=("Segoe UI", 11, "bold"), bg=_PANEL_WHITE).pack()
        
        bf = tk.Frame(dlg, bg=_PANEL_WHITE, pady=12); bf.pack(fill=tk.X, side=tk.BOTTOM)
        tk.Button(bf, text="Cancel", command=lambda: (result.set(False), dlg.destroy()), bg=_PRIMARY_BLUE, fg="#FFF", font=("Segoe UI", 10, "bold"), relief=tk.FLAT, width=12, cursor="hand2").pack(side=tk.RIGHT, padx=(10, 20))
        tk.Button(bf, text="Yes, Copy", command=lambda: (result.set(True), dlg.destroy()), bg="#E5E7EB", fg=_DARK_TEXT, font=("Segoe UI", 10), relief=tk.FLAT, width=12, cursor="hand2").pack(side=tk.RIGHT)
        
        dlg.deiconify(); dlg.wait_window()
        return result.get()

    def _on_copy_table(self) -> None:
        self.btn_copy_table.configure(state=tk.DISABLED)
        try:
            active = self._get_active_branch()
            if not active or active == "__validation_report__":
                messagebox.showinfo("Action Required", "Switch to a branch tab (CSN, CLC, or Pune) to copy.")
                return
            if active == "General":
                messagebox.showinfo(
                    "Validation Only",
                    "The General tab is for validation review only.\n"
                    "Table copy is available for Skoda tabs (CSN, CLC, Pune) only."
                )
                return
            df = self.branch_dfs.get(active)
            if df is None or df.empty: return

            # ── 1. STRICT BLOCK: UNMATCHED BE / CHECKLISTS ──
            if self.validation_results:
                be_col = "BOE NO." if "BOE NO." in df.columns else "BE No" if "BE No" in df.columns else None
                if be_col:
                    unmatched = [
                        str(r.get(be_col, "")).strip() 
                        for _, r in df.iterrows() 
                        if self.validation_results.get(str(r.get(be_col, "")).strip(), {}).get("match_method") == "No match"
                    ]
                    if unmatched:
                        messagebox.showerror(
                            "Cannot Copy", 
                            f"The following BE rows are missing a matching Checklist:\n\n{', '.join(unmatched)}\n\nPlease upload the missing Checklists or manually fix the BE numbers first."
                        )
                        return

            # ── Compulsory Fields Hard Block ──
            compulsory_cols = [
                "CHA JOB NO.", "Job No", "SUPPLIER", "Supplier", 
                "INVOICE NO.", "Invoice no.", "HAWB NO", "MBL NO/HBL NO", "BL No", 
                "HAWB DT.", "BL DATE", "BOE NO.", "BE No", "BOE DATE", "BE Dt.", 
                "ASS. VALUE", "Ass. Value", "SCRIPTS UTILISED VALUE", "Debit Duty", 
                "DUTY AMT. AS PAR BOE/CHECK LIST", "Duty", "TOTAL DUTY AMT (CASH)", "Total Amt"
            ]
            cols_to_check = [c for c in compulsory_cols if c in df.columns]
            missing_fields = []
            
            for idx, row in df.iterrows():
                for c in cols_to_check:
                    val = str(row[c]).strip()
                    if not val or val.lower() == "nan":
                        be_val = str(row.get("BOE NO.", row.get("BE No", f"Row {idx+1}"))).strip()
                        missing_fields.append(f"• BE {be_val} ➔ Missing '{c}'")

            if missing_fields:
                msg = "Cannot copy table. These required fields are missing:\n\n"
                msg += "\n".join(missing_fields[:8])
                if len(missing_fields) > 8:
                    msg += f"\n\n(+ {len(missing_fields)-8} more issues)"
                messagebox.showerror("Missing Data", msg)
                return

            # Collect Validation Issues
            issues = []
            if self.validation_results:
                be_col = "BOE NO." if "BOE NO." in df.columns else "BE No" if "BE No" in df.columns else None
                if be_col:
                    for _, row in df.iterrows():
                        be_no = str(row.get(be_col, "")).strip()
                        vr = self.validation_results.get(be_no, {})
                        st = vr.get("status", "")
                        if st not in ("HIGH RISK", "WARNING"): continue
                        diffs = []
                        if vr.get("ass_diff", 0) != 0: diffs.append(f"Assessable Value Diff: {vr['ass_diff']}")
                        if vr.get("duty_diff", 0) != 0: diffs.append(f"Duty Amount Diff: {vr['duty_diff']}")
                        if vr.get("foregone_diff", 0) != 0: diffs.append(f"Scripts vs Foregone Diff: {vr['foregone_diff']}")
                        for r in vr.get("remarks", []): diffs.append(r)
                        if not diffs: diffs.append(f"{st} issue found")
                        issues.append(f"BE {be_no} / Job {vr.get('job_no','')}\n" + "\n".join(f"  - {d}" for d in diffs))

            if issues:
                has_high = any("HIGH RISK" in self.validation_results.get(be, {}).get("status", "") for be in [str(r.get("BOE NO." if "BOE NO." in df.columns else "BE No", "")).strip() for _, r in df.iterrows()])
                severity = "HIGH RISK" if has_high else "WARNING"
                if not self._show_confirm(severity, "\n\n".join(issues)): return

            tbl_html = generate_html_table(df, exclude_cols={"Validation", "_lvl"})
            copy_html_to_clipboard(tbl_html)
            self._flash_copy_success()
        finally:
            self.btn_copy_table.configure(state=tk.NORMAL)

def main() -> None:
    root = tk.Tk()
    BECopyParserApp(root)
    root.mainloop()

if __name__ == "__main__":
    main()