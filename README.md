# AO Compare Tool

A lightweight Python script for comparing two Excel files exported from SAP Analysis for Office (AO).

It detects key columns (or uses ones you specify), identifies differences between rows, and optionally exports the results to an Excel file.

---

## 🧰 Features

- Automatic detection of header rows and logical renaming of unnamed columns
- Custom key support via `keys=...`
- Optionally limit row reading for performance with `maxline=...`
- Exports results to Excel (`comp_result.xlsx` or your custom file)
- Fully command-line driven
- Compatible with SAP AO-exported `.xlsx` structures

---

## 🚀 Usage

```bash
python CompareReports.py [options]
