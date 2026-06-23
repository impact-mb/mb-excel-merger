# Magic Bus Excel Merge Cleaner

Streamlit app to upload multiple Excel files, validate matching headers, remove Power BI footer/filter rows, merge data, and download a processed Excel file.

## Login

Default local login:

- Login ID: `admin`
- Password: `MagicBus@2026`

For deployment, change credentials in Streamlit Secrets:

```toml
APP_USERNAME = "your_login_id"
APP_PASSWORD = "your_secure_password"
```

## Run locally

```bash
pip install -r requirements.txt
streamlit run app.py
```

## Suggested GitHub repository name

`magicbus-excel-merge-cleaner`
