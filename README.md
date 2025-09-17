# AO Compare Tool

**AO Compare Tool** is a lightweight Python script for comparing two Excel files exported from SAP Analysis for Office (AO). It automatically detects header rows, identifies unique keys, compares matching rows, and outputs differences to the console or an Excel file.

---

## 🔧 Features

- Automatic header row detection
- Renames AO-style `Unnamed:` columns as `PreviousColumn (Text)`
- Supports both manual and automatic key detection
- Command-line configuration
- Optional Excel output

---

## 🚀 Usage

```
python CompareReports.py [options]
```

### Options:

| Option                | Description |
|-----------------------|-------------|
| `keys=col1,col2,...`  | Use specified columns as key |
| `maxline=500`         | Read only the first N rows |
| `base=filename.xlsx`  | Specify the base Excel file |
| `compare=filename.xlsx` | Specify the comparison Excel file |
| `exc=filename.xlsx`   | Write results to this Excel file |
| `exc=none`            | Disable Excel export |
| `debug=1`             | Print header, types, and first 3 rows |
| `--help`              | Show this help message |

---

## 🧪 Examples

```
python CompareReports.py base=a.xlsx compare=b.xlsx keys="Reference,DOC NUM"
python CompareReports.py base=a.xlsx compare=b.xlsx maxline=500 debug=1
python CompareReports.py exc=none
```

---

## 📂 Output

- **Console:**
  - Differences in rows with matching keys
  - Rows unique to the base file
  - Rows unique to the compare file

- **Excel (default: `comp_result.xlsx`):**
  - Sheet `Differences` — detailed value differences
  - Sheet `OnlyInBase` — rows only in base
  - Sheet `OnlyInCompare` — rows only in compare



## 📄 License

MIT License
