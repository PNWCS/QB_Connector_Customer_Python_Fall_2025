# Customer Connector

## Setup Project
Once you forked and cloned the repo, run:
```bash
poetry install
```
to install dependencies.
Then write code in the src/ folder.


## Build


Building as Executable
To build the project as a standalone .exe:

Install dependencies (including PyInstaller):

poetry install
Build the executable:

```bash
poetry run pyinstaller --onefile --name customers --hidden-import win32timezone --hidden-import win32com.client build_exe.py
```

The executable will be created in the dist folder.

The --hidden-import flags ensure PyInstaller includes the Windows COM dependencies needed for QuickBooks integration.

Running the Executable
After building, launch the CLI directly from Command Prompt:

Change into the dist directory (or reference the full path):
cd dist
Run the executable with the same arguments the Python entry point expects:

```bash
customers.exe --workbook C:\path\to\company_data.xlsx --output C:\path\to\report.json
```

If you omit --output, the report defaults to customer_report.json in the current directory.You can also invoke it without cd by using the absolute path, e.g.:

```bash
C:\Users\BoyaA\Desktop\QB_Connector_Customer_Python_Fall_2025\dist\customers.exe --workbook company_data.xlsx
```

## Example Output
``` bash
{
  "status": "success",
  "timestamp": "2025-12-03T19:24:32.422075+00:00",
  "added_customers": [
    {
      "record_id": "35",
      "name": "aboya",
      "source": "excel"
    }
  ],
  "conflicts": [
    {
      "record_id": "14",
      "excel_name": "xyz",
      "qb_name": "tyu",
      "reason": "data_mismatch"
    },
    {
      "record_id": "6",
      "excel_name": null,
      "qb_name": "DOLLY",
      "reason": "missing_in_excel"
    },
    {
      "record_id": "19",
      "excel_name": null,
      "qb_name": "sunny",
      "reason": "missing_in_excel"
    }
  ],
  "same_customers": 2,
  "error": null
}
```
