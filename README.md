# MB Excel Merger

Secure Streamlit app to validate, clean, and merge multiple Excel files.

## Login Users

| Region | Login ID | Password |
|---|---|---|
| North | north | North@2026 |
| East | east | East@2026 |
| West | west | West@2026 |
| South | south | South@2026 |

## Run Locally

```bash
pip install -r requirements.txt
streamlit run app.py
```

## Deploy on Streamlit Cloud

Add these values in Streamlit Cloud > App Settings > Secrets:

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
