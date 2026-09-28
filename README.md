# Excel Batch Tool
Batch-process Excel workbooks from the command line: merge multiple files, clean messy data, deduplicate rows, and build grouped summaries. No more manual copy-paste.
## Features
- Merge multiple `.xlsx` files into one
- Clean data: trim whitespace, normalize phone-number columns
- Deduplicate by key column
- Group-by aggregation (sum / count)
## Requirements
- Python 3.8+
- pandas, openpyxl (`pip install pandas openpyxl`)
## Usage
Run `python excel_tool.py` and follow the prompts: pick an operation, point it at your files, get the result.
## Tested
A 6-row sample order sheet was merged and summarized correctly (total ¥1304).
