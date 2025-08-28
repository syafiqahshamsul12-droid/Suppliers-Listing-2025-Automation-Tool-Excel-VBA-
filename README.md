# Suppliers-Listing-2025-Automation-Tool-Excel-VBA-

This Excel VBA project streamlines Accounts Payable (AP) data entry, PDF renaming, and recurring autofill logic. It significantly reduces manual workload, speeds up reporting, and improves data accuracy for invoice processing.

---

## 📝 Project Overview

During my internship, I improvised an internal AP file to:
- Automatically rename invoice PDFs according to supplier patterns.
- Auto-populate invoice-related data based on past entries by supplier and PO number.
- Reduce manual effort and streamline repetitive tasks.
- Allow users to process hundreds of invoices efficiently with minimal errors.

Estimated impact: Manually entering data that previously took hours can now be completed in **10–15 minutes** for most entries, leaving only a few items for manual input.

---

## ⚙️ Tech Stack

- Microsoft Excel + VBA  
- Excel Macros  
- Built-in Excel Formulas  
- Outlook Integration (optional for future enhancements)  

---

## 📌 Instructions for Use

**Purpose:** Process multiple invoice entries efficiently using smart macros and recurring autofill logic.

**Before You Start:**
1. Save all invoice PDFs in the same folder as this Excel file.  
2. Name folders clearly; avoid sensitive names (e.g., `Hotel Corpo 060825` → `HC - 060825`).  
3. Enter `N` in Column A for every invoice row you want to process.  
4. Click the buttons **in sequence (top to bottom)** to avoid errors.  
5. The file is reusable — macros detect where to start each time.

---

## 🔹 Workflow & Features

| Step | Action / Function | Description (User-Friendly) | Button Needed? |
|------|-----------------|----------------------------|----------------|
| 1 | Enter "N" in Column A | Start by typing `N` in Column A for each invoice. Columns like Currency, Cost Code, etc., will be auto-filled. | 🚫 No button |
| 2 | `RenameHotelCorpo()`, `RenameHolidayTours()` | Auto-renames invoice PDFs inside the folder based on supplier pattern. Example: `Facture_F_001_4829414_SBM.pdf` → `F_001_4829414.pdf`. Currently supports Hotel Corpo and Holiday Tours. | 🔘 Run |
| 3 | `AutoKeyInFileName` | Auto-fills Column D with the PDF names in the folder, starting from the first empty row. | 🔘 Run |
| 4 | Manual Supplier Selection | Use formulas in the `Reference` sheet to select supplier names. | 🚫 No button |
| 5 | Manual Data Entry | Fill in remaining blank fields (except grey cells). Description (Column V) can be left blank if needed. | Manual |
| 6 | `AutoFill` | Auto-populates Person-In-Charge, Account Code, Department/Project, Supplier Code, Asset Code & Due Date based on past entries. | 🔘 Run |
| 7 | `DetectAssetTypeAndAmount` (Optional) | Scans invoice descriptions to detect Asset Type (IT, Equipment, etc.) and Amount. | 🔘 Run |
| 8 | Final Validation | Check for duplicates, ensure all required fields are filled, and verify consistent formats. | 🚫 No button |

---

## 💡 Fun Fact
Green cells in the Excel sheet are designed to help the brain remember and focus better while processing data. 🧠

---

## 🚀 Future Improvements
- Fully automate grouped email generation with Outlook.  
- Add error handling for missing or invalid data.  
- Implement user access restrictions for sensitive columns.  

---

## ⚠️ Disclaimer
- Save this file as `.xlsm` to enable macros.  
- Do not edit VBA scripts without guidance. Consult before making changes.  
- No confidential data is included in this repository — safe for public sharing.  

---

## 🤝 Credits
Developed by **Nurul Syafiqah** during internship at **SBM Offshore Finance Department**, Malaysia. Special thanks to the Finance Ops Team for guidance and feedback.
