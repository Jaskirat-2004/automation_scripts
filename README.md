```markdown
# ⚙️ Excel Automation Scripts

A collection of advanced **Python automation scripts** built to streamline Excel-based reporting workflows.  
Developed during real-world automation projects, these scripts transform manual Excel processes into fully automated data pipelines using **pandas**, **DuckDB**, **xlwings**, and **win32com**.

---

## 🚀 Overview

This repository contains multiple end-to-end automations used across live reporting setups.  
Each script independently handles a specific workflow — such as hygiene reports, half-hourly summaries, or performance dashboards — but all follow a consistent and reusable pipeline.

**Core capabilities include:**
- 📥 Reading raw data dumps (Excel/CSV or network paths)  
- 🧮 Cleaning, transforming, and restructuring data  
- 📊 Writing processed data into Excel templates  
- 🔁 Auto-filling formulas and maintaining formatting  
- 💾 Saving final reports automatically (`_OUTPUT` versions)  
- ⚠️ Logging and error-handling for each update step  

---

## 🧠 Workflow Structure

All automation scripts are built around a standard three-stage flow:

### 1️⃣ Data Preparation
- Load raw dumps using `pandas` or SQL-style queries with **DuckDB**  
- Select and reorder columns, clean prefixes (e.g. remove `MKOC`)  
- Convert date/time columns and format consistently  

### 2️⃣ Excel Integration
- Open `.xlsb` or `.xlsx` templates through **xlwings**  
- Identify the correct paste position dynamically  
- Write cleaned DataFrames into multiple sheets  
- Apply formulas using `.AutoFill()` for scalable updates  

### 3️⃣ Output & Logging
- Save output as `_OUTPUT.xlsb`  
- Print progress logs for each step  
- Ensure proper Excel app closure to prevent COM lockups  

---

## 🧩 Tech Stack

| Tool / Library | Purpose |
|-----------------|----------|
| **Python 3.x** | Core language for scripting |
| **pandas** | Data cleaning, transformation, and manipulation |
| **DuckDB** | Fast SQL querying for Excel/CSV data |
| **xlwings** | Write/read Excel data, formula autofill |
| **win32com** | COM-based automation and legacy Excel handling |
| **os / pathlib** | File path and directory management |

---

## 📂 Repository Structure

```

/Excel_Automation_Scripts
│
├── Acko Hygiene Report.py
├── Zepto Half Hourly.py
├── Meesho SS Chat.py
│
├── /Dumps
│   ├── Raw data files used by the scripts
│
├── /Templates
│   ├── Excel templates (.xlsb / .xlsx)
│
└── README.md

````

Each script targets a different report but follows the same automation principles — making them modular, maintainable, and reusable.

---

## ⚙️ Usage

### 🔧 Installation
Install the required libraries:
```bash
pip install pandas xlwings pywin32 duckdb
````

### ▶️ Running a Script

Run any individual script from the command line or VS Code terminal:

```bash
python "Script Name.py"
```

### 🕒 Scheduling (Optional)

For full automation, integrate with **Windows Task Scheduler** to run daily or weekly at specific times.

---

## 🧭 Best Practices

* Keep template and output files separate to avoid overwriting.
* Always use **raw strings** (`r"path"`) for file paths.
* Close Excel apps cleanly in every script (`app.quit()`).
* Maintain consistent sheet names and cell reference points.
* Use `_OUTPUT` naming for finalized reports.

---

## 📘 Future Plans

* Create a unified **controller script** to trigger multiple reports sequentially
* Add **logging modules** for better traceability
* Integrate **email automation** for report distribution
* Build a **dashboard interface** to trigger scripts manually or on schedule

---

## 👨‍💻 Author

**Jaskirat**
Python Developer | Excel Automation | Data Analytics

> *"Automating the repetitive — mastering the productive."*

---

## 🏁 Summary

This project showcases the power of combining **Python**, **Excel**, and **SQL-style querying** to eliminate repetitive reporting tasks.
It’s a continuously evolving toolkit aimed at making business reporting faster, smarter, and error-free.

---

```

---
