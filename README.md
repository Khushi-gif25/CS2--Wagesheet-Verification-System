# ⚖️ Wagesheet Verification System

An automated employee-wise payroll reconciliation and statutory compliance verification system that compares vendor wagesheets with EPFO PF ECR and ESIC ECR documents to identify discrepancies, missing records, employee name mismatches, and contribution amount differences.

## 📌 Project Overview

Manual verification of vendor wagesheets against statutory compliance documents is time-consuming and prone to human errors.

This project automates the verification process by extracting employee information from Excel wagesheets and official EPFO and ESIC PDF documents, matching employee records, and identifying discrepancies.

The system generates a detailed Excel audit report and provides an interactive Streamlit dashboard for easy monitoring and compliance analysis.

## 🚀 Key Features

- **Automated Document Extraction:** Extract employee records from Excel wagesheets and EPFO/ESIC PDF documents.
- **Intelligent Name Matching:** Handles spelling variations, different word orders, abbreviations, initials, and common prefixes.
- **PF Verification:** Compares vendor PF deductions against the Employee Share (EE Share) in EPFO ECR.
- **ESIC Verification:** Compares vendor ESIC deductions against the IP Contribution in ESIC ECR.
- **Discrepancy Detection:** Identifies name mismatches, amount mismatches, missing employees, duplicate identifiers, and invalid identifiers.
- **Dynamic Column Detection:** Supports different wagesheet formats through column name aliases and header detection.
- **Adjustable Matching Threshold:** Configure name similarity from 70% to 95%, with 85% as the default.
- **Interactive Dashboard:** Displays employee-wise verification, PF summary, ESIC summary, and discrepancy findings.
- **Excel Report Generation:** Downloads a professional Excel report containing four detailed sheets.
- **Sample Dataset:** Includes a built-in sample dataset for quick testing.

## 🛠️ Technologies Used

- Python
- Streamlit
- Pandas
- OpenPyXL
- PDFPlumber
- Regular Expressions
- SequenceMatcher (Fuzzy Name Matching)

## ⚙️ How It Works

1. Upload the vendor wagesheet in Excel/CSV format.
2. Upload the official EPFO ECR PDF.
3. Upload the official ESIC ECR PDF.
4. Extract employee names, UAN, IP numbers, and contribution amounts.
5. Normalize employee identifiers and names.
6. Match employees independently against EPFO and ESIC records.
7. Compare contribution amounts and identify discrepancies.
8. Display verification results on the dashboard.
9. Download the complete Excel audit report.

## 📊 Verification Status Categories

| Status | Description |
|---|---|
| MATCHED | Identifier, employee name, and contribution amount match |
| NAME MISMATCH | Employee name differs from the ECR record |
| AMOUNT MISMATCH | Contribution amount differs from the ECR |
| MISSING | Employee identifier not found in ECR |
| DUPLICATE IN ECR | Duplicate identifier found in ECR |
| INVALID IDENTIFIER | Required UAN or IP number is missing |

## 📁 Excel Report Structure

The application generates a downloadable Excel report with four sheets:

**1. Employee-wise Verification**
- Employee Name
- UAN and IP Number
- Vendor PF and ESIC amounts
- EPFO and ESIC contribution amounts
- Verification status
- Amount differences and reasons

**2. PF Summary**
- Total employees checked
- Matched employees
- Name mismatches
- Amount mismatches
- Missing employees
- Other discrepancies

**3. ESIC Summary**
- Total employees checked
- Matched employees
- Name mismatches
- Amount mismatches
- Missing employees
- Other discrepancies

**4. Overall Summary**
- Total employees processed
- PF and ESIC matched records
- Total discrepancies
- Fully compliant employees
- Overall compliance rate
- Detailed discrepancy findings

## 💻 Installation & Setup

### 1. Clone the Repository

```bash
git clone https://github.com/Khushi-gif25/CS2--Wagesheet-Verification-System.git
```

### 2. Navigate to Project Directory

```bash
cd CS2--Wagesheet-Verification-System
```

### 3. Create a Virtual Environment (Optional)

```bash
python -m venv venv
```

Activate it on Windows:

```bash
venv\Scripts\activate
```

### 4. Install Dependencies

```bash
pip install -r requirements.txt
```

### 5. Run the Application

```bash
streamlit run app.py
```

The application will open in your browser.

## 📂 Input Documents

The application supports:

- Vendor Wagesheet: `.xlsx`, `.xls`, `.xlsm`, `.csv`
- EPFO ECR: `.pdf`
- ESIC ECR: `.pdf`

A built-in sample dataset is also available for testing the application.

## 🎯 Project Objective

The primary objective is to automate payroll verification and statutory compliance auditing by reducing manual reconciliation efforts, improving accuracy, identifying discrepancies, and generating structured audit reports.

This solution helps organizations simplify vendor payroll verification and improve the efficiency of PF and ESIC compliance checks.

## 🔮 Future Enhancements

- Integration with additional payroll management systems.
- Advanced OCR support for scanned PDF documents.
- Automated email alerts for critical discrepancies.
- Historical audit tracking and comparison.
- Cloud deployment for centralized access.
- Enhanced audit analytics and visualization.

---

**Project:** Case Study 2 – Wagesheet Verification System  
**Category:** Payroll Automation | Data Processing | Compliance Auditing | Python Automation
