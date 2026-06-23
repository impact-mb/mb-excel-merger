# MB Excel Merger

Secure Streamlit application to validate, clean, and merge multiple `.csv` or `.xls` files.

## Why this app

- Standardizes the data merge process across regions.
- Reduces manual Excel copy-paste work.
- Checks that all uploaded files have matching headers before merging.
- Removes unwanted Power BI footer/filter rows such as `Total` and `Applied filters:`.
- Creates a clean Excel output with file-wise processing summary.

## Supported input files

Only these formats are supported:

- `.csv`
- `.xls`

The final processed output is downloaded as `.xlsx`.

## Run locally

```bash
pip install -r requirements.txt
streamlit run app.py
```

## Streamlit Cloud secrets

Add credentials in Streamlit Cloud > App Settings > Secrets.

```toml
[users.north]
password = "North@2026"
region = "North"

[users.east]
password = "East@2026"
region = "East"

[users.west]
password = "West@2026"
region = "West"

[users.south]
password = "South@2026"
region = "South"
```

## File structure

```text
MB Excel Merger/
├── app.py
├── magicbus_logo.png
├── requirements.txt
├── README.md
└── .streamlit/
    └── secrets.toml.example
```
