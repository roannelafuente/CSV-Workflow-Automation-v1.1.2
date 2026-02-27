# CSV-Workflow-Automation-v1.1.2

## 📖 Description
Incremental update to the CSV Workflow Automation Tool, focusing on normalization, branding, and robust error handling.  
This release does not introduce major new features but instead strengthens reliability, maintainability, and professional polish while retaining all existing functionality from v1.1.1.

---

## 📌 Disclaimer
This project is a portfolio demonstration built entirely with synthetic/dummy data.  
While the workbook structure and formatting are inspired by typical engineering workflows, all headers, values, and examples have been replaced with generic placeholders.  
No proprietary intellectual property, client data, or company‑specific conventions are included.  
Its sole purpose is to showcase automation techniques, reproducible workflows, and technical skills in Python, Tkinter, pandas, OpenPyXL, and xlwings.

---

## 🚀 Changes in v1.1.2
- **Normalization**
  - Improved consistency in `C1_MARK` lookups by normalizing values (e.g., `"1.0"` → `"1"`) specifically for **color mapping in wafermap visualization**.
- **Branding**
  - Added custom `.ico` icon (`sprout.ico`) for professional GUI branding, with fallback handling if the file is missing.
- **Error Handling**
  - Integrated centralized `ErrorLogger`:
    - Auto‑creates timestamped error log files in a dedicated `/logs` folder.
    - Cleans up logs older than 30 days automatically.
    - Ensures consistent error capture across all modules with dual reporting (GUI + log files).
- **Architecture**
  - Refactored into a modular, multi‑class design for cleaner code and easier maintenance.
  - Added helper functions for normalization and header lookup.
  - Optimized bulk operations for faster CSV → Excel conversions.
- **Documentation**
  - Consistent docstrings and portfolio‑ready structure.

---

## 🛠️ Tech Stack
- Python (automation & GUI)  
- Tkinter (user interface)  
- pandas (bulk CSV handling)  
- OpenPyXL (Excel file handling)  
- xlwings (pivot tables & wafermap generation)  
- CSV (data parsing)  

---

## 📂 Sample Files
- Input: [`DEMO_CSV_SAMPLE.csv`](https://github.com/roannelafuente/CSV-Workflow-Automation/blob/main/DEMO_CSV_SAMPLE.csv)  
- Output: [`DEMO_CSV_SAMPLE.xlsx`](https://github.com/roannelafuente/CSV-Workflow-Automation/blob/main/DEMO_CSV_SAMPLE.xlsx)  
- Dashboard Screenshot: `CSV Workflow Automation Dashboard v1.1.2.png`  

These files are dummy inputs included for demonstration purposes only. They are synthetic examples and do not contain any client or company data. Their purpose is to allow recruiters and collaborators to quickly test the workflow and visualize the results.

---

## 📸 Screenshots
### GUI Dashboard
![GUI Dashboard](https://github.com/roannelafuente/CSV-Workflow-Automation-v1.1.2/blob/main/CSV%20Workflow%20Automation%20Dashboard%20v1.1.2.png)

---

## 🌟 Impact
- **Consistency:** Normalized `C1_MARK` values ensure accurate wafermap color mapping and pivot filtering.  
- **Professionalism:** Custom `.ico` icon and GUI branding elevate portfolio presentation.  
- **Reliability:** Centralized error logging captures issues across all modules, with automatic cleanup for long‑term maintainability.  
- **Maintainability:** Multi‑class refactor makes the codebase easier to extend, debug, and showcase as a portfolio project.  
- **Efficiency:** Optimized bulk operations reduce runtime for large CSV → Excel conversions.  
- **Transparency:** Scrollable status box and dual error reporting (GUI + logs) improve user confidence and workflow clarity.  

---

## 📦 Download
The latest release (**v1.1.2**) with normalization, branding, robust error handling, and architecture improvements is available here:  
[➡️ Download CSV Workflow Automation Tool v1.1.2](https://github.com/roannelafuente/CSV-Workflow-Automation-v1.1.2/releases/tag/v1.1.2)

▶️ **Usage**:  
Run the `.exe` to launch the dashboard and explore the features.

---

## 👩‍💻 Author
**Rose Anne Lafuente**  
Licensed Electronics Engineer | Product Engineer II | Python Automation  
GitHub: [@roannelafuente](https://github.com/roannelafuente)  
LinkedIn: [Rose Anne Lafuente](www.linkedin.com/in/rose-anne-lafuente)
