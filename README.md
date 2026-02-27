# Auto Data Pipeline – Full Excel Automation

## Description

Auto Data Pipeline is a fully automated system that:  

- Cleans raw Excel files  
- Removes duplicates and empty rows  
- Keeps only specified columns  
- Generates formatted reports by Region and Product  
- Applies professional Excel styling and conditional formatting  
- Archives processed files  

It is **cross-platform** (Windows, Linux, macOS) and configurable via a simple `config.json` file.  



##  Features

- 🔹 **Cross-platform:** Works on Windows, Linux, and macOS  
- 🔹 **Configurable:** Folder paths, thresholds, and columns via `config.json`  
- 🔹 **Excel Automation:** Cleans, analyzes, and formats files automatically  
- 🔹 **Logging:** Tracks every step in `automation.log`  
- 🔹 **Archiving:** Automatically moves processed files to preserve originals  



##  Installation

1️⃣ Clone the repository:

```bash
git clone https://github.com/DIANEDWEIRI/auto-data-pipeline.git

2️⃣ Install required packages:

pip install pandas openpyxl
📂 Usage
1️⃣ Configure Settings

Edit config.json to match your folder structure and preferences:

{
  "base_folder": "Desktop/AutoDataPipeline",
  "input_folder": "input/raw_excel",
  "cleaned_folder": "cleaned/cleaned_excel",
  "reports_folder": "reports/formatted_reports",
  "logs_folder": "logs",
  "thresholds": {
    "green": 200,
    "red": 50
  },
  "columns_to_keep": ["Region", "Product", "TotalPrice"]
}

"base_folder" → Main folder where everything lives

"input_folder" → Raw Excel files folder

"cleaned_folder" → Folder for cleaned Excel files

"reports_folder" → Folder for generated reports

"logs_folder" → Folder for logging

"thresholds" → Conditional formatting thresholds

"columns_to_keep" → Columns to keep in the cleaned data

2️⃣ Run the Pipeline
python pipeline.py

The script will process all .xlsx files in the input folder.

Cleaned files are saved in cleaned_folder.

Formatted reports are saved in reports_folder.

Original files are archived with PROCESSED_ prefix.

Logs are saved in logs/automation.log.