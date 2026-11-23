"""
Wagesheet vs ECR Compliance Verification Application
===================================================
A production-grade, generalized compliance verification system to cross-check
Vendor Wagesheets against EPFO ECR (EE Share) and ESIC ECR (IP Contribution) documents.

Key Architectural Components:
1. Dynamic Document Extractors (Excel wagesheets, EPFO ECR PDFs, ESIC ECR PDFs)
2. Intelligent Name Normalization & Multi-tier Matching Engine
3. Statutory Compliance Verification Processor (Independent PF & ESIC audits)
4. Executive-grade Multi-sheet Excel Audit Report Generator (openpyxl)
5. Interactive Web Dashboard (Streamlit) & Headless CLI Interface
"""

import os
import sys
import re
import io
import argparse
from difflib import SequenceMatcher
import pandas as pd
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
import pdfplumber

# Try importing streamlit runtime
try:
    from streamlit.runtime import exists as st_runtime_exists
except ImportError:
    st_runtime_exists = lambda: False

# =============================================================================
# 1. HELPER UTILITIES & DATA CLEANING
# =============================================================================

def clean_identifier(val):
    """
    Cleans identification numbers (UAN, IP Number).
    Handles floats, scientific notation, trailing decimals, spaces, and non-digits.
    """
    if pd.isna(val) or val is None:
        return ""
    if isinstance(val, (int, float)):
        try:
            return str(int(round(val))).strip()
        except Exception:
            pass
    s = str(val).strip()
    if s.endswith('.0'):
        s = s[:-2]
    # Remove all non-digits
    return re.sub(r'\D', '', s)

def clean_amount(val):
    """
    Cleans currency and numerical amount fields.
    Removes currency symbols (₹, Rs.), commas, whitespace, and handles nulls.
    """
    if pd.isna(val) or val is None:
        return 0.0
    if isinstance(val, (int, float)):
        return float(val)
    s = str(val).strip()
    s = re.sub(r'[^\d.-]', '', s)
    try:
        return float(s)
    except Exception:
        return 0.0

# =============================================================================
# 2. INTELLIGENT NAME MATCHING LOGIC
# =============================================================================

def normalize_name(name):
    """
    Normalizes employee names by:
    - Converting to uppercase
    - Stripping punctuation, hyphens, and dots
    - Collapsing redundant whitespace
    - Removing common standalone honorifics/prefixes (MR, MS, MRS, SMT, SH, SHRI, DR)
    """
    if not name or pd.isna(name):
        return ""
    s = str(name).upper().strip()
    s = re.sub(r'[^\w\s]', ' ', s)
    words = s.split()
    prefixes = {'MR', 'MS', 'MRS', 'SMT', 'SH', 'SHRI', 'DR'}
    if words and words[0] in prefixes and len(words) > 1:
        words = words[1:]
    return " ".join(words)

def compare_names(name1, name2, threshold=0.85):
    """
    Multi-tier intelligent name matching algorithm:
    Tier 1: Exact normalized string match
    Tier 2: Token-set / Word-order permutation match (e.g., 'ANIL KUMAR' vs 'KUMAR ANIL')
    Tier 3: Initial / Abbreviation matching (e.g., 'RAMESH C GUPTA' vs 'RAMESH CHANDRA GUPTA')
    Tier 4: Fuzzy SequenceMatcher similarity with first-name safeguard
    
    Returns:
        (is_match: bool, score: float, explanation: str)
    """
    norm1 = normalize_name(name1)
    norm2 = normalize_name(name2)
    
    if not norm1 or not norm2:
        return False, 0.0, "Empty name identifier"
    
    # Tier 1: Exact Match
    if norm1 == norm2:
        return True, 1.0, "Exact match"
    
    words1 = norm1.split()
    words2 = norm2.split()
    
    # Tier 2: Word-order Invariance
    if sorted(words1) == sorted(words2):
        return True, 1.0, "Word order variation"
        
    # Tier 3: Abbreviation / Initial Expansion
    if len(words1) == len(words2) and len(words1) >= 2:
        exact_tokens = 0
        initial_tokens = 0
        for w1, w2 in zip(words1, words2):
            if w1 == w2:
                exact_tokens += 1
            elif (len(w1) == 1 and w2.startswith(w1)) or (len(w2) == 1 and w1.startswith(w2)):
                initial_tokens += 1
        if exact_tokens + initial_tokens == len(words1) and exact_tokens >= 1:
            return True, 0.90, "Initial abbreviation match"
            
    # Tier 4: Fuzzy Sequence Matching
    sorted_s1 = " ".join(sorted(words1))
    sorted_s2 = " ".join(sorted(words2))
    
    ratio_direct = SequenceMatcher(None, norm1, norm2).ratio()
    ratio_sorted = SequenceMatcher(None, sorted_s1, sorted_s2).ratio()
    score = max(ratio_direct, ratio_sorted)
    
    # First-name safeguard: prevents fundamentally distinct names with common surnames
    # (e.g., 'DINESH PRASAD' vs 'DEVENDRA PRASAD' or 'REKHA SHARMA' vs 'RENU SHARMA')
    # from incorrectly matching.
    w1_first = words1[0]
    w2_first = words2[0]
    first_ratio = SequenceMatcher(None, w1_first, w2_first).ratio()
    
    if score >= threshold:
        if len(words1) > 1 and len(words2) > 1 and first_ratio < 0.60:
            return False, round(score, 3), f"Different first name ({w1_first} vs {w2_first})"
        return True, round(score, 3), f"Fuzzy match (score: {score:.2f})"
    else:
        return False, round(score, 3), f"Similarity {score:.2f} below threshold {threshold}"

# =============================================================================
# 3. DYNAMIC DATA EXTRACTORS
# =============================================================================

def extract_wagesheet_excel(file_source):
    """
    Dynamically identifies and extracts employee records from any vendor wagesheet.
    Scans leading rows to locate the header dynamically regardless of metadata/title rows.
    Dynamically identifies columns using semantic aliases.
    """
    if isinstance(file_source, bytes):
        file_source = io.BytesIO(file_source)
        
    df_raw = pd.read_excel(file_source, header=None)
    header_row_idx = None
    
    # Scan first 20 rows to detect header
    for idx, row in df_raw.iloc[:20].iterrows():
        row_str = " ".join([str(x).lower() for x in row.dropna()])
        has_id = any(k in row_str for k in ['uan', 'ip', 'pf', 'esi'])
        has_name = any(k in row_str for k in ['name', 'emp', 'worker', 'employee'])
        if has_id and has_name:
            header_row_idx = idx
            break
            
    if hasattr(file_source, 'seek'):
        file_source.seek(0)
        
    if header_row_idx is not None:
        df = pd.read_excel(file_source, header=header_row_idx)
    else:
        df = pd.read_excel(file_source)
        
    # Map column headers dynamically
    col_mapping = {}
    for col in df.columns:
        c_clean = str(col).lower().strip()
        if any(k in c_clean for k in ['emp name', 'employee name', 'worker name', 'member name', 'name']) and 'col_name' not in col_mapping:
            col_mapping['col_name'] = col
        elif any(k in c_clean for k in ['uan no', 'uan number', 'uan', 'pf uan', 'epf uan']) and 'col_uan' not in col_mapping:
            col_mapping['col_uan'] = col
        elif any(k in c_clean for k in ['ip number', 'ip no', 'ip_number', 'ip', 'insurance no', 'esic ip', 'esi no', 'esic no']) and 'col_ip' not in col_mapping:
            col_mapping['col_ip'] = col
        elif any(k in c_clean for k in ['agency', 'vendor', 'contractor', 'company', 'agency name']) and 'col_agency' not in col_mapping:
            col_mapping['col_agency'] = col
        elif any(k in c_clean for k in ['state', 'location', 'branch', 'region']) and 'col_state' not in col_mapping:
            col_mapping['col_state'] = col
        elif any(k in c_clean for k in ['esic @', 'esic@', 'esic', 'esi', 'esi @', 'esi@', 'ip contribution', 'esic contribution']) and 'col_esic' not in col_mapping:
            col_mapping['col_esic'] = col
        elif any(k in c_clean for k in ['pf @', 'pf@', 'epf', 'pf', 'ee share', 'ee contribution', 'pf contribution']) and 'col_pf' not in col_mapping:
            col_mapping['col_pf'] = col

    records = []
    for _, row in df.iterrows():
        name_val = row.get(col_mapping.get('col_name'))
        name = str(name_val).strip() if pd.notna(name_val) else ""
        if not name or name.lower() in ['nan', 'none', 'total', 'grand total', 'sub total']:
            continue
            
        uan = clean_identifier(row.get(col_mapping.get('col_uan'), ''))
        ip = clean_identifier(row.get(col_mapping.get('col_ip'), ''))
        agency = str(row.get(col_mapping.get('col_agency'), '')).strip() if 'col_agency' in col_mapping and pd.notna(row[col_mapping['col_agency']]) else ""
        state = str(row.get(col_mapping.get('col_state'), '')).strip() if 'col_state' in col_mapping and pd.notna(row[col_mapping['col_state']]) else ""
        esic_amt = clean_amount(row.get(col_mapping.get('col_esic'), 0))
        pf_amt = clean_amount(row.get(col_mapping.get('col_pf'), 0))
        
        records.append({
            'name': name,
            'uan': uan,
            'ip_number': ip,
            'agency': agency,
            'state': state,
            'esic_amount': esic_amt,
            'pf_amount': pf_amt
        })
    return records

def extract_pf_ecr(pdf_source):
    """
    Extracts member records from EPFO Electronic Challan cum Return (ECR) PDF.
    Extracts UAN, Name as per ECR, and EE Share (Employee Contribution).
    Handles multi-page documents, table layout shifts, and large page counts.
    """
    records = []
    metadata = {}
    
    if isinstance(pdf_source, bytes):
        pdf_file = io.BytesIO(pdf_source)
    else:
        pdf_file = pdf_source
        
    with pdfplumber.open(pdf_file) as pdf:
        for page_idx, page in enumerate(pdf.pages):
            text = page.extract_text()
            if not text:
                continue
                
            # Extract header metadata if present on first page
            if page_idx == 0:
                est_m = re.search(r'Name of Establishment\s+(.+)', text)
                if est_m:
                    metadata['establishment_name'] = est_m.group(1).split('\n')[0].strip()
                month_m = re.search(r'Wage Month\s+([A-Z0-9-]+)', text)
                if month_m:
                    metadata['wage_month'] = month_m.group(1).strip()
                members_m = re.search(r'Total Members\s+(\d+)', text)
                if members_m:
                    metadata['total_members'] = int(members_m.group(1))
                    
            lines = text.split('\n')
            for line in lines:
                line = line.strip()
                if not line:
                    continue
                # Line format in official EPFO ECR:
                # SlNo UAN Name_ECR Name_Repo Gross EPF EPS EDLI EE EPS ER NCP
                match = re.match(r'^(\d+)\s+(\d{12})\s+(.*?)\s+([\d,]+\.?\d*)\s+([\d,]+\.?\d*)\s+([\d,]+\.?\d*)\s+([\d,]+\.?\d*)\s+([\d,]+\.?\d*)\s+([\d,]+\.?\d*)\s+([\d,]+\.?\d*)\s+(\d+)$', line)
                if match:
                    sl_no, uan, name_part, gross, epf, eps, edli, ee, eps_contrib, er, ncp = match.groups()
                    ee_val = float(ee.replace(',', ''))
                    
                    # Name column in EPFO ECR contains "Name as per ECR" and "Name as per UAN Repository"
                    name_words = name_part.strip().split()
                    if len(name_words) % 2 == 0:
                        half = len(name_words) // 2
                        first_half = " ".join(name_words[:half])
                        second_half = " ".join(name_words[half:])
                        clean_name = first_half if first_half == second_half else first_half
                    else:
                        clean_name = name_part
                        
                    records.append({
                        'sl_no': int(sl_no),
                        'uan': uan,
                        'name': clean_name.strip(),
                        'ee_share': ee_val,
                        'gross_wages': float(gross.replace(',', '')),
                        'epf_wages': float(epf.replace(',', '')),
                        'eps_share': float(eps_contrib.replace(',', '')),
                        'er_share': float(er.replace(',', '')),
                        'page': page_idx + 1
                    })
    return records, metadata

def extract_esic_ecr(pdf_source):
    """
    Extracts insured person records from ESIC Contribution History PDF.
    Extracts IP Number, IP Name, and IP Contribution (Employee Share @ 0.75%).
    Handles multi-page documents and table layouts.
    """
    records = []
    metadata = {}
    
    if isinstance(pdf_source, bytes):
        pdf_file = io.BytesIO(pdf_source)
    else:
        pdf_file = pdf_source
        
    with pdfplumber.open(pdf_file) as pdf:
        for page_idx, page in enumerate(pdf.pages):
            text = page.extract_text()
            if not text:
                continue
                
            # Extract header metadata if present on first page
            if page_idx == 0:
                header_m = re.search(r'Contribution History Of\s+(\d+)\s+for\s+([A-Za-z0-9]+)', text)
                if header_m:
                    metadata['employer_code'] = header_m.group(1).strip()
                    metadata['wage_month'] = header_m.group(2).strip()
                tot_ip_m = re.search(r'Total IP Contribution\s+.*?\n\s*([\d,]+\.?\d*)', text)
                if tot_ip_m:
                    metadata['total_ip_contribution'] = float(tot_ip_m.group(1).replace(',', ''))
                    
            lines = text.split('\n')
            for line in lines:
                line = line.strip()
                if not line:
                    continue
                # Line format in ESIC Contribution History:
                # SNo. IsDisable IPNumber IPName NoOfDays TotalWages IPContribution Reason
                match = re.match(r'^(\d+)\s+([^\s]+)\s+(\d{10,17})\s+(.+?)\s+(\d+)\s+([\d,]+\.?\d*)\s+([\d,]+\.?\d*)(?:\s+(.*))?$', line)
                if match:
                    sno, is_disable, ip_num, name, days, total_wages, ip_contrib, reason = match.groups()
                    contrib_val = float(ip_contrib.replace(',', ''))
                    records.append({
                        'sno': int(sno),
                        'ip_number': ip_num,
                        'name': name.strip(),
                        'days': int(days),
                        'total_wages': float(total_wages.replace(',', '')),
                        'ip_contribution': contrib_val,
                        'reason': reason.strip() if reason else '-',
                        'page': page_idx + 1
                    })
    return records, metadata

# =============================================================================
# 4. COMPLIANCE VERIFICATION PROCESSOR
# =============================================================================

def verify_compliance(wagesheet_records, pf_records, esic_records, name_threshold=0.85):
    """
    Performs independent statutory compliance audits for PF and ESIC for every employee.
    
    Verdicts for PF and ESIC:
    - MATCHED: Identifier found, name matches, amount matches submitted ECR.
    - NAME MISMATCH: Identifier found, but name differs between Wagesheet and ECR.
    - AMOUNT MISMATCH: Identifier and name match, but submitted amount differs.
    - MISSING: Employee identifier cannot be found in the corresponding ECR.
    - DUPLICATE IN ECR: Identifier appears multiple times in ECR (flagged for review).
    - INVALID IDENTIFIER: Identifier is missing or blank in wagesheet.
    """
    # Index PF records by UAN (handling duplicate detection)
    pf_by_uan = {}
    for r in pf_records:
        uan = r['uan']
        pf_by_uan.setdefault(uan, []).append(r)
        
    # Index ESIC records by IP Number (handling duplicate detection)
    esic_by_ip = {}
    for r in esic_records:
        ip = r['ip_number']
        esic_by_ip.setdefault(ip, []).append(r)
        
    results = []
    for emp in wagesheet_records:
        emp_name = emp['name']
        uan = emp['uan']
        ip = emp['ip_number']
        agency = emp['agency']
        ws_pf = emp['pf_amount']
        ws_esic = emp['esic_amount']
        
        # ---------------------------------------------------------------------
        # Independent PF Audit
        # ---------------------------------------------------------------------
        if not uan:
            pf_status = "INVALID IDENTIFIER"
            pf_ecr_ee = None
            pf_diff = ws_pf
            pf_reason = "UAN missing in wagesheet"
        elif uan not in pf_by_uan:
            pf_status = "MISSING"
            pf_ecr_ee = None
            pf_diff = ws_pf
            pf_reason = f"UAN {uan} not found in EPFO ECR"
        elif len(pf_by_uan[uan]) > 1:
            pf_status = "DUPLICATE IN ECR"
            pf_ecr_ee = pf_by_uan[uan][0]['ee_share']
            pf_diff = ws_pf - pf_ecr_ee
            pf_reason = f"Duplicate UAN found {len(pf_by_uan[uan])} times in EPFO ECR. Manual review required."
        else:
            pf_record = pf_by_uan[uan][0]
            pf_ecr_name = pf_record['name']
            pf_ecr_ee = pf_record['ee_share']
            pf_diff = round(ws_pf - pf_ecr_ee, 2)
            
            # Name Verification
            is_match, score, n_reason = compare_names(emp_name, pf_ecr_name, threshold=name_threshold)
            if not is_match:
                pf_status = "NAME MISMATCH"
                pf_reason = f"Name mismatch: Wagesheet has '{emp_name}', EPFO ECR has '{pf_ecr_name}'"
            elif abs(pf_diff) > 0.01:
                pf_status = "AMOUNT MISMATCH"
                pf_reason = f"Amount mismatch: Wagesheet PF is ₹{ws_pf:,.2f}, EPFO EE Share is ₹{pf_ecr_ee:,.2f} (Diff: ₹{pf_diff:,.2f})"
            else:
                pf_status = "MATCHED"
                pf_reason = f"Verified: UAN, Name, and EE Share match (₹{ws_pf:,.2f})"
                
        # ---------------------------------------------------------------------
        # Independent ESIC Audit
        # ---------------------------------------------------------------------
        if not ip:
            esic_status = "INVALID IDENTIFIER"
            esic_ecr_ip = None
            esic_diff = ws_esic
            esic_reason = "IP Number missing in wagesheet"
        elif ip not in esic_by_ip:
            esic_status = "MISSING"
            esic_ecr_ip = None
            esic_diff = ws_esic
            esic_reason = f"IP Number {ip} not found in ESIC ECR"
        elif len(esic_by_ip[ip]) > 1:
            esic_status = "DUPLICATE IN ECR"
            esic_ecr_ip = esic_by_ip[ip][0]['ip_contribution']
            esic_diff = ws_esic - esic_ecr_ip
            esic_reason = f"Duplicate IP Number found {len(esic_by_ip[ip])} times in ESIC ECR. Manual review required."
        else:
            esic_record = esic_by_ip[ip][0]
            esic_ecr_name = esic_record['name']
            esic_ecr_ip = esic_record['ip_contribution']
            esic_diff = round(ws_esic - esic_ecr_ip, 2)
            
            # Name Verification
            is_match, score, n_reason = compare_names(emp_name, esic_ecr_name, threshold=name_threshold)
            if not is_match:
                esic_status = "NAME MISMATCH"
                esic_reason = f"Name mismatch: Wagesheet has '{emp_name}', ESIC ECR has '{esic_ecr_name}'"
            elif abs(esic_diff) > 0.01:
                esic_status = "AMOUNT MISMATCH"
                esic_reason = f"Amount mismatch: Wagesheet ESIC is ₹{ws_esic:,.2f}, ESIC IP Contrib is ₹{esic_ecr_ip:,.2f} (Diff: ₹{esic_diff:,.2f})"
            else:
                esic_status = "MATCHED"
                esic_reason = f"Verified: IP Number, Name, and IP Contribution match (₹{ws_esic:,.2f})"
                
        results.append({
            'Employee Name': emp_name,
            'UAN': uan,
            'IP Number': ip,
            'Agency': agency,
            'Wagesheet PF Amount': ws_pf,
            'PF ECR EE Share': pf_ecr_ee,
            'PF Status': pf_status,
            'PF Difference': pf_diff,
            'PF Reason': pf_reason,
            'Wagesheet ESIC Amount': ws_esic,
            'ESIC ECR IP Contribution': esic_ecr_ip,
            'ESIC Status': esic_status,
            'ESIC Difference': esic_diff,
            'ESIC Reason': esic_reason
        })
        
    return results

# =============================================================================
# 5. EXCEL AUDIT REPORT GENERATOR (4 SHEETS)
# =============================================================================

def generate_excel_report(results, output_target=None):
    """
    Generates a downloadable Excel audit report with 4 professional sheets:
    - Sheet 1: Employee-wise Verification
    - Sheet 2: PF Summary
    - Sheet 3: ESIC Summary
    - Sheet 4: Overall Summary
    
    If output_target is None, returns bytes buffer.
    """
    wb = openpyxl.Workbook()
    wb.remove(wb.active)  # Remove blank default sheet
    
    font_family = "Segoe UI"
    header_font = Font(name=font_family, size=11, bold=True, color="FFFFFF")
    title_font = Font(name=font_family, size=14, bold=True, color="1F4E79")
    subtitle_font = Font(name=font_family, size=10, italic=True, color="595959")
    regular_font = Font(name=font_family, size=10)
    bold_font = Font(name=font_family, size=10, bold=True)
    
    header_fill_main = PatternFill(start_color="1F4E79", end_color="1F4E79", fill_type="solid")
    header_fill_pf = PatternFill(start_color="203764", end_color="203764", fill_type="solid")
    header_fill_esic = PatternFill(start_color="375623", end_color="375623", fill_type="solid")
    summary_header_fill = PatternFill(start_color="2F5597", end_color="2F5597", fill_type="solid")
    
    status_styles = {
        "MATCHED": {
            "fill": PatternFill(start_color="E2EFDA", end_color="E2EFDA", fill_type="solid"),
            "font": Font(name=font_family, size=10, bold=True, color="375623")
        },
        "NAME MISMATCH": {
            "fill": PatternFill(start_color="FFF2CC", end_color="FFF2CC", fill_type="solid"),
            "font": Font(name=font_family, size=10, bold=True, color="B25E00")
        },
        "AMOUNT MISMATCH": {
            "fill": PatternFill(start_color="FCE4D6", end_color="FCE4D6", fill_type="solid"),
            "font": Font(name=font_family, size=10, bold=True, color="C65911")
        },
        "MISSING": {
            "fill": PatternFill(start_color="F8CBAD", end_color="F8CBAD", fill_type="solid"),
            "font": Font(name=font_family, size=10, bold=True, color="C00000")
        },
        "DUPLICATE IN ECR": {
            "fill": PatternFill(start_color="FFD9D9", end_color="FFD9D9", fill_type="solid"),
            "font": Font(name=font_family, size=10, bold=True, color="9C0006")
        },
        "INVALID IDENTIFIER": {
            "fill": PatternFill(start_color="F2DCDB", end_color="F2DCDB", fill_type="solid"),
            "font": Font(name=font_family, size=10, bold=True, color="9C0006")
        }
    }
    
    thin_border = Border(
        left=Side(style='thin', color='D9D9D9'),
        right=Side(style='thin', color='D9D9D9'),
        top=Side(style='thin', color='D9D9D9'),
        bottom=Side(style='thin', color='D9D9D9')
    )
    
    # -------------------------------------------------------------------------
    # Sheet 1: Employee-wise Verification
    # -------------------------------------------------------------------------
    ws1 = wb.create_sheet(title="Employee-wise Verification")
    ws1.views.sheetView[0].showGridLines = True
    
    headers_sheet1 = [
        "Employee Name", "UAN", "IP Number", "Agency",
        "Wagesheet PF Amount", "PF ECR EE Share", "PF Status", "PF Difference", "PF Reason",
        "Wagesheet ESIC Amount", "ESIC ECR IP Contribution", "ESIC Status", "ESIC Difference", "ESIC Reason"
    ]
    
    ws1.merge_cells("A1:N1")
    ws1["A1"] = "WAGESHEET VS ECR COMPLIANCE VERIFICATION REPORT"
    ws1["A1"].font = title_font
    ws1["A1"].alignment = Alignment(vertical="center")
    ws1.row_dimensions[1].height = 25
    
    ws1.merge_cells("A2:N2")
    ws1["A2"] = "Verification of Vendor Wagesheet deductions against EPFO ECR (EE Share) and ESIC ECR (IP Contribution)"
    ws1["A2"].font = subtitle_font
    ws1.row_dimensions[2].height = 18
    
    ws1.row_dimensions[4].height = 28
    for col_idx, header in enumerate(headers_sheet1, 1):
        cell = ws1.cell(row=4, column=col_idx, value=header)
        cell.font = header_font
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        if "PF" in header:
            cell.fill = header_fill_pf
        elif "ESIC" in header:
            cell.fill = header_fill_esic
        else:
            cell.fill = header_fill_main
            
    current_row = 5
    for item in results:
        ws1.row_dimensions[current_row].height = 22
        
        ws1.cell(row=current_row, column=1, value=item["Employee Name"]).alignment = Alignment(vertical="center")
        ws1.cell(row=current_row, column=2, value=str(item["UAN"])).alignment = Alignment(horizontal="center", vertical="center")
        ws1.cell(row=current_row, column=3, value=str(item["IP Number"])).alignment = Alignment(horizontal="center", vertical="center")
        ws1.cell(row=current_row, column=4, value=item["Agency"]).alignment = Alignment(vertical="center")
        
        # PF
        c5 = ws1.cell(row=current_row, column=5, value=item["Wagesheet PF Amount"])
        c5.number_format = '#,##0.00'
        c5.alignment = Alignment(horizontal="right", vertical="center")
        
        c6 = ws1.cell(row=current_row, column=6, value=item["PF ECR EE Share"] if item["PF ECR EE Share"] is not None else "-")
        if item["PF ECR EE Share"] is not None:
            c6.number_format = '#,##0.00'
        c6.alignment = Alignment(horizontal="right", vertical="center")
        
        c7 = ws1.cell(row=current_row, column=7, value=item["PF Status"])
        c7.alignment = Alignment(horizontal="center", vertical="center")
        if item["PF Status"] in status_styles:
            c7.fill = status_styles[item["PF Status"]]["fill"]
            c7.font = status_styles[item["PF Status"]]["font"]
            
        c8 = ws1.cell(row=current_row, column=8, value=item["PF Difference"])
        c8.number_format = '#,##0.00'
        c8.alignment = Alignment(horizontal="right", vertical="center")
        
        c9 = ws1.cell(row=current_row, column=9, value=item["PF Reason"])
        c9.alignment = Alignment(vertical="center")
        
        # ESIC
        c10 = ws1.cell(row=current_row, column=10, value=item["Wagesheet ESIC Amount"])
        c10.number_format = '#,##0.00'
        c10.alignment = Alignment(horizontal="right", vertical="center")
        
        c11 = ws1.cell(row=current_row, column=11, value=item["ESIC ECR IP Contribution"] if item["ESIC ECR IP Contribution"] is not None else "-")
        if item["ESIC ECR IP Contribution"] is not None:
            c11.number_format = '#,##0.00'
        c11.alignment = Alignment(horizontal="right", vertical="center")
        
        c12 = ws1.cell(row=current_row, column=12, value=item["ESIC Status"])
        c12.alignment = Alignment(horizontal="center", vertical="center")
        if item["ESIC Status"] in status_styles:
            c12.fill = status_styles[item["ESIC Status"]]["fill"]
            c12.font = status_styles[item["ESIC Status"]]["font"]
            
        c13 = ws1.cell(row=current_row, column=13, value=item["ESIC Difference"])
        c13.number_format = '#,##0.00'
        c13.alignment = Alignment(horizontal="right", vertical="center")
        
        c14 = ws1.cell(row=current_row, column=14, value=item["ESIC Reason"])
        c14.alignment = Alignment(vertical="center")
        
        for c in range(1, 15):
            cell = ws1.cell(row=current_row, column=c)
            cell.border = thin_border
            if c not in [7, 12]:
                cell.font = regular_font
                
        current_row += 1

    for col in ws1.columns:
        max_len = 0
        col_letter = get_column_letter(col[0].column)
        for cell in col:
            if cell.row < 4:
                continue
            val = str(cell.value or '')
            if len(val) > max_len:
                max_len = len(val)
        ws1.column_dimensions[col_letter].width = max(max_len + 3, 12)

    # -------------------------------------------------------------------------
    # Sheet 2: PF Summary
    # -------------------------------------------------------------------------
    ws2 = wb.create_sheet(title="PF Summary")
    ws2.views.sheetView[0].showGridLines = True
    
    total_emp = len(results)
    pf_matched = sum(1 for r in results if r["PF Status"] == "MATCHED")
    pf_name_mismatch = sum(1 for r in results if r["PF Status"] == "NAME MISMATCH")
    pf_amount_mismatch = sum(1 for r in results if r["PF Status"] == "AMOUNT MISMATCH")
    pf_missing = sum(1 for r in results if r["PF Status"] == "MISSING")
    pf_other = total_emp - (pf_matched + pf_name_mismatch + pf_amount_mismatch + pf_missing)
    
    ws2.merge_cells("A1:C1")
    ws2["A1"] = "EPFO COMPLIANCE VERIFICATION SUMMARY"
    ws2["A1"].font = title_font
    
    ws2.row_dimensions[3].height = 24
    ws2.cell(row=3, column=1, value="Metric").font = header_font
    ws2.cell(row=3, column=1).fill = header_fill_pf
    ws2.cell(row=3, column=2, value="Count").font = header_font
    ws2.cell(row=3, column=2).fill = header_fill_pf
    ws2.cell(row=3, column=3, value="% of Total").font = header_font
    ws2.cell(row=3, column=3).fill = header_fill_pf
    
    pf_rows = [
        ("Total employees checked", total_emp, 1.0, Font(name=font_family, size=10, bold=True)),
        ("Matched employees", pf_matched, pf_matched/total_emp if total_emp else 0, status_styles["MATCHED"]["font"]),
        ("Name mismatches", pf_name_mismatch, pf_name_mismatch/total_emp if total_emp else 0, status_styles["NAME MISMATCH"]["font"]),
        ("Amount mismatches", pf_amount_mismatch, pf_amount_mismatch/total_emp if total_emp else 0, status_styles["AMOUNT MISMATCH"]["font"]),
        ("Missing employees", pf_missing, pf_missing/total_emp if total_emp else 0, status_styles["MISSING"]["font"]),
    ]
    if pf_other > 0:
        pf_rows.append(("Other Discrepancies / Duplicates", pf_other, pf_other/total_emp if total_emp else 0, bold_font))
        
    for idx, (m_label, m_val, m_pct, m_font) in enumerate(pf_rows, start=4):
        ws2.row_dimensions[idx].height = 20
        c1 = ws2.cell(row=idx, column=1, value=m_label)
        c1.font = regular_font
        c1.border = thin_border
        
        c2 = ws2.cell(row=idx, column=2, value=m_val)
        c2.font = m_font
        c2.alignment = Alignment(horizontal="right")
        c2.border = thin_border
        
        c3 = ws2.cell(row=idx, column=3, value=m_pct)
        c3.number_format = '0.0%'
        c3.font = regular_font
        c3.alignment = Alignment(horizontal="right")
        c3.border = thin_border

    ws2.column_dimensions['A'].width = 32
    ws2.column_dimensions['B'].width = 15
    ws2.column_dimensions['C'].width = 15

    # -------------------------------------------------------------------------
    # Sheet 3: ESIC Summary
    # -------------------------------------------------------------------------
    ws3 = wb.create_sheet(title="ESIC Summary")
    ws3.views.sheetView[0].showGridLines = True
    
    esic_matched = sum(1 for r in results if r["ESIC Status"] == "MATCHED")
    esic_name_mismatch = sum(1 for r in results if r["ESIC Status"] == "NAME MISMATCH")
    esic_amount_mismatch = sum(1 for r in results if r["ESIC Status"] == "AMOUNT MISMATCH")
    esic_missing = sum(1 for r in results if r["ESIC Status"] == "MISSING")
    esic_other = total_emp - (esic_matched + esic_name_mismatch + esic_amount_mismatch + esic_missing)
    
    ws3.merge_cells("A1:C1")
    ws3["A1"] = "ESIC COMPLIANCE VERIFICATION SUMMARY"
    ws3["A1"].font = title_font
    
    ws3.row_dimensions[3].height = 24
    ws3.cell(row=3, column=1, value="Metric").font = header_font
    ws3.cell(row=3, column=1).fill = header_fill_esic
    ws3.cell(row=3, column=2, value="Count").font = header_font
    ws3.cell(row=3, column=2).fill = header_fill_esic
    ws3.cell(row=3, column=3, value="% of Total").font = header_font
    ws3.cell(row=3, column=3).fill = header_fill_esic
    
    esic_rows = [
        ("Total employees checked", total_emp, 1.0, Font(name=font_family, size=10, bold=True)),
        ("Matched employees", esic_matched, esic_matched/total_emp if total_emp else 0, status_styles["MATCHED"]["font"]),
        ("Name mismatches", esic_name_mismatch, esic_name_mismatch/total_emp if total_emp else 0, status_styles["NAME MISMATCH"]["font"]),
        ("Amount mismatches", esic_amount_mismatch, esic_amount_mismatch/total_emp if total_emp else 0, status_styles["AMOUNT MISMATCH"]["font"]),
        ("Missing employees", esic_missing, esic_missing/total_emp if total_emp else 0, status_styles["MISSING"]["font"]),
    ]
    if esic_other > 0:
        esic_rows.append(("Other Discrepancies / Duplicates", esic_other, esic_other/total_emp if total_emp else 0, bold_font))
        
    for idx, (m_label, m_val, m_pct, m_font) in enumerate(esic_rows, start=4):
        ws3.row_dimensions[idx].height = 20
        c1 = ws3.cell(row=idx, column=1, value=m_label)
        c1.font = regular_font
        c1.border = thin_border
        
        c2 = ws3.cell(row=idx, column=2, value=m_val)
        c2.font = m_font
        c2.alignment = Alignment(horizontal="right")
        c2.border = thin_border
        
        c3 = ws3.cell(row=idx, column=3, value=m_pct)
        c3.number_format = '0.0%'
        c3.font = regular_font
        c3.alignment = Alignment(horizontal="right")
        c3.border = thin_border

    ws3.column_dimensions['A'].width = 32
    ws3.column_dimensions['B'].width = 15
    ws3.column_dimensions['C'].width = 15

    # -------------------------------------------------------------------------
    # Sheet 4: Overall Summary
    # -------------------------------------------------------------------------
    ws4 = wb.create_sheet(title="Overall Summary")
    ws4.views.sheetView[0].showGridLines = True
    
    pf_discrepancies = total_emp - pf_matched
    esic_discrepancies = total_emp - esic_matched
    fully_compliant = sum(1 for r in results if r["PF Status"] == "MATCHED" and r["ESIC Status"] == "MATCHED")
    compliance_pct = (fully_compliant / total_emp * 100) if total_emp else 0.0
    overall_status = "FULLY COMPLIANT" if pf_discrepancies == 0 and esic_discrepancies == 0 else f"NON-COMPLIANT ({pf_discrepancies} PF discrepancies, {esic_discrepancies} ESIC discrepancies)"
    
    ws4.merge_cells("A1:D1")
    ws4["A1"] = "OVERALL STATUTORY COMPLIANCE AUDIT REPORT"
    ws4["A1"].font = title_font
    
    ws4.row_dimensions[3].height = 24
    ws4.cell(row=3, column=1, value="Summary Parameter").font = header_font
    ws4.cell(row=3, column=1).fill = summary_header_fill
    ws4.cell(row=3, column=2, value="Audit Value").font = header_font
    ws4.cell(row=3, column=2).fill = summary_header_fill
    
    overall_kpis = [
        ("Total employees processed", total_emp),
        ("Total PF matched", pf_matched),
        ("Total ESIC matched", esic_matched),
        ("Total PF discrepancies", pf_discrepancies),
        ("Total ESIC discrepancies", esic_discrepancies),
        ("Fully Compliant Employees (Both PF & ESIC)", fully_compliant),
        ("Overall Compliance Rate", f"{compliance_pct:.1f}%"),
        ("Overall compliance status", overall_status)
    ]
    
    for idx, (param, val) in enumerate(overall_kpis, start=4):
        ws4.row_dimensions[idx].height = 20
        c1 = ws4.cell(row=idx, column=1, value=param)
        c1.font = bold_font if idx in [4, 9, 10, 11] else regular_font
        c1.border = thin_border
        
        c2 = ws4.cell(row=idx, column=2, value=val)
        c2.font = bold_font if idx in [4, 9, 10, 11] else regular_font
        if param == "Overall compliance status":
            c2.font = Font(name=font_family, size=10, bold=True, color="C00000" if "NON" in str(val) else "375623")
        c2.alignment = Alignment(horizontal="left" if isinstance(val, str) else "right")
        c2.border = thin_border
        
    # Flagged Employees Detailed Section
    start_flagged = 14
    ws4.cell(row=start_flagged, column=1, value="DISCREPANCY DETAILS / FLAGGED EMPLOYEES").font = Font(name=font_family, size=11, bold=True, color="C00000")
    
    flag_headers = ["Employee Name", "Identifier", "Type", "Status", "Difference (₹)", "Specific Audit Finding / Reason"]
    ws4.row_dimensions[start_flagged + 1].height = 22
    for c_idx, fh in enumerate(flag_headers, 1):
        c = ws4.cell(row=start_flagged + 1, column=c_idx, value=fh)
        c.font = header_font
        c.fill = PatternFill(start_color="833C0C", end_color="833C0C", fill_type="solid")
        c.alignment = Alignment(horizontal="center", vertical="center")
        
    f_row = start_flagged + 2
    for r in results:
        # PF discrepancy
        if r["PF Status"] != "MATCHED":
            ws4.row_dimensions[f_row].height = 20
            ws4.cell(row=f_row, column=1, value=r["Employee Name"]).font = regular_font
            ws4.cell(row=f_row, column=2, value=str(r["UAN"])).alignment = Alignment(horizontal="center")
            ws4.cell(row=f_row, column=3, value="PF").alignment = Alignment(horizontal="center")
            st_cell = ws4.cell(row=f_row, column=4, value=r["PF Status"])
            st_cell.alignment = Alignment(horizontal="center")
            if r["PF Status"] in status_styles:
                st_cell.fill = status_styles[r["PF Status"]]["fill"]
                st_cell.font = status_styles[r["PF Status"]]["font"]
            diff_cell = ws4.cell(row=f_row, column=5, value=r["PF Difference"])
            diff_cell.number_format = '#,##0.00'
            diff_cell.alignment = Alignment(horizontal="right")
            ws4.cell(row=f_row, column=6, value=r["PF Reason"])
            for col_i in range(1, 7):
                ws4.cell(row=f_row, column=col_i).border = thin_border
            f_row += 1
            
        # ESIC discrepancy
        if r["ESIC Status"] != "MATCHED":
            ws4.row_dimensions[f_row].height = 20
            ws4.cell(row=f_row, column=1, value=r["Employee Name"]).font = regular_font
            ws4.cell(row=f_row, column=2, value=str(r["IP Number"])).alignment = Alignment(horizontal="center")
            ws4.cell(row=f_row, column=3, value="ESIC").alignment = Alignment(horizontal="center")
            st_cell = ws4.cell(row=f_row, column=4, value=r["ESIC Status"])
            st_cell.alignment = Alignment(horizontal="center")
            if r["ESIC Status"] in status_styles:
                st_cell.fill = status_styles[r["ESIC Status"]]["fill"]
                st_cell.font = status_styles[r["ESIC Status"]]["font"]
            diff_cell = ws4.cell(row=f_row, column=5, value=r["ESIC Difference"])
            diff_cell.number_format = '#,##0.00'
            diff_cell.alignment = Alignment(horizontal="right")
            ws4.cell(row=f_row, column=6, value=r["ESIC Reason"])
            for col_i in range(1, 7):
                ws4.cell(row=f_row, column=col_i).border = thin_border
            f_row += 1

    ws4.column_dimensions['A'].width = 25
    ws4.column_dimensions['B'].width = 16
    ws4.column_dimensions['C'].width = 10
    ws4.column_dimensions['D'].width = 20
    ws4.column_dimensions['E'].width = 16
    ws4.column_dimensions['F'].width = 65

    if output_target:
        wb.save(output_target)
        return output_target
    else:
        buffer = io.BytesIO()
        wb.save(buffer)
        buffer.seek(0)
        return buffer

# =============================================================================
# 6. STREAMLIT WEB INTERFACE
# =============================================================================

def run_streamlit_app():
    import streamlit as st
    
    st.set_page_config(
        page_title="Wagesheet vs ECR Compliance Verification",
        page_icon="⚖️",
        layout="wide",
        initial_sidebar_state="expanded"
    )
    
    # Custom CSS for executive dashboard styling
    st.markdown("""
    <style>
    .main-header {
        font-size: 1.8rem;
        font-weight: 700;
        color: #1F4E79;
        margin-bottom: 0.2rem;
    }
    .sub-header {
        font-size: 1.0rem;
        color: #555555;
        margin-bottom: 1.5rem;
    }
    .metric-card {
        background-color: #F8F9FA;
        border-radius: 8px;
        padding: 16px;
        border-left: 5px solid #1F4E79;
        box-shadow: 0 1px 3px rgba(0,0,0,0.1);
    }
    .badge-matched {
        background-color: #E2EFDA;
        color: #375623;
        padding: 4px 10px;
        border-radius: 4px;
        font-weight: 600;
        font-size: 0.85rem;
    }
    .badge-name-mismatch {
        background-color: #FFF2CC;
        color: #B25E00;
        padding: 4px 10px;
        border-radius: 4px;
        font-weight: 600;
        font-size: 0.85rem;
    }
    .badge-amount-mismatch {
        background-color: #FCE4D6;
        color: #C65911;
        padding: 4px 10px;
        border-radius: 4px;
        font-weight: 600;
        font-size: 0.85rem;
    }
    .badge-missing {
        background-color: #F8CBAD;
        color: #C00000;
        padding: 4px 10px;
        border-radius: 4px;
        font-weight: 600;
        font-size: 0.85rem;
    }
    </style>
    """, unsafe_allow_html=True)

    st.markdown('<div class="main-header">⚖️ Wagesheet vs ECR Compliance Verification</div>', unsafe_allow_html=True)
    st.markdown('<div class="sub-header">Automated statutory cross-audit verifying vendor wagesheet deductions against official EPFO and ESIC Electronic Challan cum Returns.</div>', unsafe_allow_html=True)
    
    # Locate sample files directory
    current_dir = os.path.dirname(os.path.abspath(__file__))
    sample_dir = os.path.join(current_dir, "sample_files")
    sample_ws_path = os.path.join(sample_dir, "Vendor_Wagesheet_AUG2025.xlsx")
    sample_pf_path = os.path.join(sample_dir, "PF_ECR_AUG-25_SAMPLE.pdf")
    sample_esic_path = os.path.join(sample_dir, "ESIC_ECR_AUG-25_SAMPLE.pdf")
    
    sample_available = (
        os.path.exists(sample_ws_path) and 
        os.path.exists(sample_pf_path) and 
        os.path.exists(sample_esic_path)
    )

    # -------------------------------------------------------------------------
    # Sidebar: Document Upload & Settings
    # -------------------------------------------------------------------------
    with st.sidebar:
        st.header("📁 Document Ingestion")
        st.write("Upload vendor documents to perform compliance verification.")
        
        # Quick Load Sample Files
        if sample_available:
            st.info("💡 **Quick Test:** Load built-in sample files to test all 6 discrepancy scenarios instantly.")
            if st.button("🚀 Load Sample Dataset (Aug 2025)", use_container_width=True):
                st.session_state["use_sample"] = True
                st.session_state["ran_verification"] = False
                st.rerun()
                
        use_sample = st.session_state.get("use_sample", False)
        
        if use_sample:
            st.success("✅ Sample files loaded (Aug 2025 dataset).")
            if st.button("Reset / Upload Custom Files"):
                st.session_state["use_sample"] = False
                st.session_state["ran_verification"] = False
                st.rerun()
                
        # File uploaders
        ws_file = None if use_sample else st.file_uploader(
            "1. Vendor Wagesheet (Excel)", 
            type=["xlsx", "xls", "xlsm", "csv"],
            help="Upload the vendor monthly wagesheet containing employee deductions."
        )
        
        pf_file = None if use_sample else st.file_uploader(
            "2. EPFO ECR PDF", 
            type=["pdf"],
            help="Upload the official Electronic Challan cum Return PDF from EPFO."
        )
        
        esic_file = None if use_sample else st.file_uploader(
            "3. ESIC ECR PDF", 
            type=["pdf"],
            help="Upload the official Contribution History PDF from ESIC."
        )
        
        st.markdown("---")
        st.subheader("⚙️ Verification Rules")
        name_thresh = st.slider("Name Similarity Threshold", min_value=0.70, max_value=0.95, value=0.85, step=0.05,
                                help="Minimum similarity ratio required to match names. Default is 0.85.")
        
        run_btn = st.button("🔍 Run Verification", type="primary", use_container_width=True)

    # -------------------------------------------------------------------------
    # Execution & Processing
    # -------------------------------------------------------------------------
    has_files = (use_sample and sample_available) or (ws_file and pf_file and esic_file)
    
    if run_btn:
        if not has_files:
            st.error("⚠️ Please upload all three documents (Vendor Wagesheet, EPFO ECR, and ESIC ECR) or load the sample files.")
            return
        st.session_state["ran_verification"] = True
        
    if st.session_state.get("ran_verification", False) and has_files:
        with st.spinner("Extracting documents and running compliance audits..."):
            try:
                # 1. Extract Wagesheet
                if use_sample:
                    ws_records = extract_wagesheet_excel(sample_ws_path)
                else:
                    ws_records = extract_wagesheet_excel(ws_file.getvalue())
                    
                # 2. Extract PF ECR
                if use_sample:
                    pf_records, pf_meta = extract_pf_ecr(sample_pf_path)
                else:
                    pf_records, pf_meta = extract_pf_ecr(pf_file.getvalue())
                    
                # 3. Extract ESIC ECR
                if use_sample:
                    esic_records, esic_meta = extract_esic_ecr(sample_esic_path)
                else:
                    esic_records, esic_meta = extract_esic_ecr(esic_file.getvalue())
                    
                # 4. Verify Compliance
                results = verify_compliance(ws_records, pf_records, esic_records, name_threshold=name_thresh)
                
            except Exception as e:
                st.error(f"❌ Error during processing: {str(e)}")
                return
                
        # ---------------------------------------------------------------------
        # KPI Metric Summary
        # ---------------------------------------------------------------------
        total_emp = len(results)
        pf_matched = sum(1 for r in results if r["PF Status"] == "MATCHED")
        pf_discrepancies = total_emp - pf_matched
        
        esic_matched = sum(1 for r in results if r["ESIC Status"] == "MATCHED")
        esic_discrepancies = total_emp - esic_matched
        
        fully_compliant = sum(1 for r in results if r["PF Status"] == "MATCHED" and r["ESIC Status"] == "MATCHED")
        compliance_rate = (fully_compliant / total_emp * 100) if total_emp else 0.0
        
        kpi_col1, kpi_col2, kpi_col3, kpi_col4 = st.columns(4)
        
        with kpi_col1:
            st.metric("Total Employees Audited", total_emp, help="Total employee records extracted from vendor wagesheet")
        with kpi_col2:
            st.metric("EPFO Matched", f"{pf_matched}/{total_emp}", f"{pf_discrepancies} discrepancies", delta_color="inverse")
        with kpi_col3:
            st.metric("ESIC Matched", f"{esic_matched}/{total_emp}", f"{esic_discrepancies} discrepancies", delta_color="inverse")
        with kpi_col4:
            st.metric("Overall Compliance Rate", f"{compliance_rate:.1f}%", 
                      "Fully Compliant" if pf_discrepancies == 0 and esic_discrepancies == 0 else "Audit Flags Detected",
                      delta_color="normal" if pf_discrepancies == 0 and esic_discrepancies == 0 else "inverse")

        st.markdown("---")
        
        # Download Excel Report Action
        excel_buffer = generate_excel_report(results)
        st.download_button(
            label="📥 Download Complete Audit Report (4-Sheet Excel)",
            data=excel_buffer,
            file_name="Wagesheet_ECR_Compliance_Audit_Report.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            type="primary"
        )
        
        st.markdown("<br>", unsafe_allow_html=True)

        # ---------------------------------------------------------------------
        # Multi-tab Audit Presentation
        # ---------------------------------------------------------------------
        tab1, tab2, tab3, tab4, tab5 = st.tabs([
            "📋 Employee-wise Verification", 
            "🏛️ EPFO Audit Summary", 
            "🏥 ESIC Audit Summary", 
            "⚠️ Discrepancy Findings",
            "📖 Architecture & Rules"
        ])
        
        # --- TAB 1: Employee-wise Verification ---
        with tab1:
            st.subheader("Employee-wise Verification Breakdown")
            
            # Filter options
            filter_choice = st.radio(
                "Filter records:", 
                ["All Employees", "Discrepancies Only (Flags)", "Matched Only", "PF Flags Only", "ESIC Flags Only"], 
                horizontal=True
            )
            
            filtered_results = results
            if filter_choice == "Discrepancies Only (Flags)":
                filtered_results = [r for r in results if r["PF Status"] != "MATCHED" or r["ESIC Status"] != "MATCHED"]
            elif filter_choice == "Matched Only":
                filtered_results = [r for r in results if r["PF Status"] == "MATCHED" and r["ESIC Status"] == "MATCHED"]
            elif filter_choice == "PF Flags Only":
                filtered_results = [r for r in results if r["PF Status"] != "MATCHED"]
            elif filter_choice == "ESIC Flags Only":
                filtered_results = [r for r in results if r["ESIC Status"] != "MATCHED"]
                
            df_display = pd.DataFrame(filtered_results)
            
            # Display formatted dataframe
            st.dataframe(
                df_display[[
                    "Employee Name", "UAN", "IP Number", "Agency",
                    "Wagesheet PF Amount", "PF ECR EE Share", "PF Status", "PF Difference", "PF Reason",
                    "Wagesheet ESIC Amount", "ESIC ECR IP Contribution", "ESIC Status", "ESIC Difference", "ESIC Reason"
                ]],
                column_config={
                    "Wagesheet PF Amount": st.column_config.NumberColumn(format="₹%.2f"),
                    "PF ECR EE Share": st.column_config.NumberColumn(format="₹%.2f"),
                    "PF Difference": st.column_config.NumberColumn(format="₹%.2f"),
                    "Wagesheet ESIC Amount": st.column_config.NumberColumn(format="₹%.2f"),
                    "ESIC ECR IP Contribution": st.column_config.NumberColumn(format="₹%.2f"),
                    "ESIC Difference": st.column_config.NumberColumn(format="₹%.2f"),
                },
                use_container_width=True,
                height=450
            )
            
        # --- TAB 2: EPFO Summary ---
        with tab2:
            st.subheader("EPFO (PF) Compliance Audit Breakdown")
            pf_name_mismatch = sum(1 for r in results if r["PF Status"] == "NAME MISMATCH")
            pf_amount_mismatch = sum(1 for r in results if r["PF Status"] == "AMOUNT MISMATCH")
            pf_missing = sum(1 for r in results if r["PF Status"] == "MISSING")
            pf_other = total_emp - (pf_matched + pf_name_mismatch + pf_amount_mismatch + pf_missing)
            
            pf_sum_data = [
                {"Metric": "Total Employees Checked", "Count": total_emp, "% of Total": f"100.0%"},
                {"Metric": "Matched Employees", "Count": pf_matched, "% of Total": f"{(pf_matched/total_emp*100):.1f}%"},
                {"Metric": "Name Mismatches", "Count": pf_name_mismatch, "% of Total": f"{(pf_name_mismatch/total_emp*100):.1f}%"},
                {"Metric": "Amount Mismatches", "Count": pf_amount_mismatch, "% of Total": f"{(pf_amount_mismatch/total_emp*100):.1f}%"},
                {"Metric": "Missing in EPFO ECR", "Count": pf_missing, "% of Total": f"{(pf_missing/total_emp*100):.1f}%"},
            ]
            if pf_other > 0:
                pf_sum_data.append({"Metric": "Other / Duplicates", "Count": pf_other, "% of Total": f"{(pf_other/total_emp*100):.1f}%"})
                
            st.table(pd.DataFrame(pf_sum_data))
            
            st.info("""
            **EPFO Audit Rules Applied:**
            - Comparison is performed strictly against the **EE (Employee Contribution)** column of the ECR.
            - EPS (Pension) and ER (Employer Share) contributions are ignored per statutory case study rules.
            - Matching identifier: **12-digit UAN**.
            """)

        # --- TAB 3: ESIC Summary ---
        with tab3:
            st.subheader("ESIC Compliance Audit Breakdown")
            esic_name_mismatch = sum(1 for r in results if r["ESIC Status"] == "NAME MISMATCH")
            esic_amount_mismatch = sum(1 for r in results if r["ESIC Status"] == "AMOUNT MISMATCH")
            esic_missing = sum(1 for r in results if r["ESIC Status"] == "MISSING")
            esic_other = total_emp - (esic_matched + esic_name_mismatch + esic_amount_mismatch + esic_missing)
            
            esic_sum_data = [
                {"Metric": "Total Employees Checked", "Count": total_emp, "% of Total": f"100.0%"},
                {"Metric": "Matched Employees", "Count": esic_matched, "% of Total": f"{(esic_matched/total_emp*100):.1f}%"},
                {"Metric": "Name Mismatches", "Count": esic_name_mismatch, "% of Total": f"{(esic_name_mismatch/total_emp*100):.1f}%"},
                {"Metric": "Amount Mismatches", "Count": esic_amount_mismatch, "% of Total": f"{(esic_amount_mismatch/total_emp*100):.1f}%"},
                {"Metric": "Missing in ESIC ECR", "Count": esic_missing, "% of Total": f"{(esic_missing/total_emp*100):.1f}%"},
            ]
            if esic_other > 0:
                esic_sum_data.append({"Metric": "Other / Duplicates", "Count": esic_other, "% of Total": f"{(esic_other/total_emp*100):.1f}%"})
                
            st.table(pd.DataFrame(esic_sum_data))
            
            st.info("""
            **ESIC Audit Rules Applied:**
            - Comparison is performed strictly against the **IP Contribution** column of the ECR.
            - Total Employer Contribution and Total Remittance columns are ignored.
            - Matching identifier: **10 to 17 digit IP Number**.
            """)

        # --- TAB 4: Discrepancy Findings ---
        with tab4:
            st.subheader("⚠️ Specific Discrepancy Findings & Audit Evidence")
            
            discrepancy_list = []
            for r in results:
                if r["PF Status"] != "MATCHED":
                    discrepancy_list.append({
                        "Employee Name": r["Employee Name"],
                        "Identifier": r["UAN"],
                        "Authority": "EPFO",
                        "Audit Verdict": r["PF Status"],
                        "Difference (₹)": f"₹{r['PF Difference']:,.2f}",
                        "Audit Finding": r["PF Reason"]
                    })
                if r["ESIC Status"] != "MATCHED":
                    discrepancy_list.append({
                        "Employee Name": r["Employee Name"],
                        "Identifier": r["IP Number"],
                        "Authority": "ESIC",
                        "Audit Verdict": r["ESIC Status"],
                        "Difference (₹)": f"₹{r['ESIC Difference']:,.2f}",
                        "Audit Finding": r["ESIC Reason"]
                    })
                    
            if discrepancy_list:
                st.dataframe(pd.DataFrame(discrepancy_list), use_container_width=True)
            else:
                st.success("🎉 No compliance discrepancies detected! 100% statutory match.")

        # --- TAB 5: Architecture & Rules ---
        with tab5:
            st.subheader("📖 Verification Logic & Methodology Documentation")
            st.markdown("""
            ### 1. Document Extraction Architecture
            - **Vendor Wagesheet (Excel):** Dynamically scans file rows for column aliases (`Name`, `UAN`, `IP Number`, `PF @ 12%`, `ESIC @ 0.75%`). Strips scientific notation and decimal formatting from identifiers.
            - **EPFO ECR (PDF):** Extracts member data across all pages. Identifies the member row and extracts the **EE (Employee Share)** column. Ignores EPS and ER columns.
            - **ESIC ECR (PDF):** Extracts insured person records across all pages. Maps the **IP Contribution** column. Ignores Employer Contribution.

            ### 2. Four Verification Verdicts
            1. **`MATCHED`:**
               - Identifier (UAN/IP) located in ECR.
               - Name verified within similarity threshold.
               - Wagesheet deduction exactly equals submitted ECR contribution.
            2. **`NAME MISMATCH`:**
               - Identifier located in ECR.
               - Name failed similarity check (e.g., 'DINESH PRASAD' vs 'DEVENDRA PRASAD' or 'REKHA SHARMA' vs 'RENU SHARMA').
            3. **`AMOUNT MISMATCH`:**
               - Identifier and name match.
               - Submitted amount differs (e.g., Wagesheet ₹1,654 vs ECR ₹1,354).
               - Exact difference calculated and reported.
            4. **`MISSING`:**
               - Identifier absent from corresponding ECR.
            5. **`DUPLICATE IN ECR`:**
               - Identifier appears multiple times in ECR; flagged for manual review rather than selecting arbitrary records.

            ### 3. Intelligent Name Matching Rules
            - **Normalization:** Uppercase conversion, punctuation stripping, whitespace collapsing.
            - **Honorific Filtering:** Strips MR, MS, MRS, SMT, SH, SHRI, DR prefixes.
            - **Word-order Invariance:** Treats 'ANIL KUMAR' and 'KUMAR ANIL' as identical.
            - **Abbreviation/Initials:** Resolves single-letter expansions (e.g., 'RAMESH C GUPTA' vs 'RAMESH CHANDRA GUPTA').
            - **First-name Safeguard:** Ensures different first names do not match even with identical surnames.
            """)
    else:
        st.info("👈 Please upload the vendor wagesheet and ECR PDF files in the sidebar, or click **'Load Sample Dataset'** and then **'Run Verification'**.")

# =============================================================================
# 7. CLI INTERFACE & MAIN ENTRYPOINT
# =============================================================================

def run_cli():
    parser = argparse.ArgumentParser(description="Wagesheet vs ECR Compliance Verification CLI")
    parser.add_argument("--wagesheet", required=True, help="Path to Vendor Wagesheet Excel file")
    parser.add_argument("--pf", required=True, help="Path to EPFO ECR PDF file")
    parser.add_argument("--esic", required=True, help="Path to ESIC ECR PDF file")
    parser.add_argument("--output", default="compliance_verification_report.xlsx", help="Output path for Excel report")
    parser.add_argument("--threshold", type=float, default=0.85, help="Name similarity threshold (default 0.85)")
    args = parser.parse_args()

    print(f"Loading Wagesheet: {args.wagesheet}")
    ws_records = extract_wagesheet_excel(args.wagesheet)
    print(f"Extracted {len(ws_records)} employees from wagesheet.")

    print(f"Loading EPFO ECR: {args.pf}")
    pf_records, pf_meta = extract_pf_ecr(args.pf)
    print(f"Extracted {len(pf_records)} members from EPFO ECR.")

    print(f"Loading ESIC ECR: {args.esic}")
    esic_records, esic_meta = extract_esic_ecr(args.esic)
    print(f"Extracted {len(esic_records)} insured persons from ESIC ECR.")

    print("Executing compliance verification...")
    results = verify_compliance(ws_records, pf_records, esic_records, name_threshold=args.threshold)

    print(f"Generating Excel report: {args.output}")
    generate_excel_report(results, output_target=args.output)
    print("Verification completed successfully!")

if __name__ == "__main__":
    # Check if CLI flag is present
    if "--cli" in sys.argv:
        sys.argv.remove("--cli")
        run_cli()
    elif not st_runtime_exists():
        # If run directly via 'python app.py', launch Streamlit
        import subprocess
        print("Launching Streamlit Web Application...")
        subprocess.run(["streamlit", "run", __file__])
    else:
        # Running under Streamlit runner
        run_streamlit_app()
