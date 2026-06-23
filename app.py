import io
import hmac
from datetime import datetime
from typing import List, Tuple

import pandas as pd
import streamlit as st

APP_NAME = "Magic Bus Excel Merge Cleaner"
LOGO_PATH = "magicbus_logo.png"

# Default login for local use. For Streamlit Cloud/GitHub deployment, use .streamlit/secrets.toml.
DEFAULT_USERNAME = "admin"
DEFAULT_PASSWORD = "MagicBus@2026"

DROP_TEXTS = {
    "total",
    "applied filters:",
}


def get_credentials() -> Tuple[str, str]:
    """Read login credentials from Streamlit secrets, otherwise use local defaults."""
    username = st.secrets.get("APP_USERNAME", DEFAULT_USERNAME)
    password = st.secrets.get("APP_PASSWORD", DEFAULT_PASSWORD)
    return str(username), str(password)


def check_login(username: str, password: str) -> bool:
    saved_username, saved_password = get_credentials()
    return hmac.compare_digest(username, saved_username) and hmac.compare_digest(password, saved_password)


def login_screen() -> None:
    st.set_page_config(page_title=f"Login | {APP_NAME}", page_icon="🚌", layout="centered")

    c1, c2, c3 = st.columns([1, 2, 1])
    with c2:
        st.image(LOGO_PATH, width=170)
        st.title("Login")
        st.caption(APP_NAME)

        with st.form("login_form"):
            username = st.text_input("Login ID")
            password = st.text_input("Password", type="password")
            submitted = st.form_submit_button("Login", type="primary", use_container_width=True)

        if submitted:
            if check_login(username, password):
                st.session_state["authenticated"] = True
                st.session_state["login_user"] = username
                st.rerun()
            else:
                st.error("Invalid Login ID or Password.")

        with st.expander("Default local login"):
            st.write(f"Login ID: `{DEFAULT_USERNAME}`")
            st.write(f"Password: `{DEFAULT_PASSWORD}`")


def logout_button() -> None:
    if st.sidebar.button("Logout"):
        st.session_state.clear()
        st.rerun()


def normalize_headers(headers: List[str]) -> List[str]:
    return [str(h).strip() for h in headers]


def read_excel_as_text(uploaded_file) -> pd.DataFrame:
    """Read first worksheet as text only."""
    return pd.read_excel(uploaded_file, dtype=str, engine="openpyxl")


def row_should_drop(row: pd.Series) -> bool:
    values = ["" if pd.isna(v) else str(v).strip() for v in row.tolist()]
    joined = " ".join(values).strip()
    joined_lower = joined.lower()

    if joined_lower in DROP_TEXTS:
        return True

    if joined_lower.startswith("applied filters:"):
        return True

    # Power BI export footer/filter rows often have only one populated cell.
    non_blank = [v for v in values if v != ""]
    if len(non_blank) == 1:
        only_value = non_blank[0].strip().lower()
        if only_value == "total" or only_value.startswith("applied filters:"):
            return True

    return False


def clean_dataframe(df: pd.DataFrame) -> Tuple[pd.DataFrame, int]:
    before = len(df)
    df = df.dropna(how="all")
    mask = df.apply(row_should_drop, axis=1)
    df = df.loc[~mask].copy()
    removed = before - len(df)
    return df, removed


def build_output_excel(df: pd.DataFrame) -> bytes:
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name="Merged_Data")
        summary = pd.DataFrame(
            {
                "Metric": ["Processed On", "Total Final Rows", "Total Columns"],
                "Value": [datetime.now().strftime("%d-%m-%Y %H:%M:%S"), len(df), len(df.columns)],
            }
        )
        summary.to_excel(writer, index=False, sheet_name="Summary")
    output.seek(0)
    return output.getvalue()


if not st.session_state.get("authenticated", False):
    login_screen()
    st.stop()

st.set_page_config(page_title=APP_NAME, page_icon="🚌", layout="wide")

with st.sidebar:
    st.image(LOGO_PATH, width=180)
    st.title("Magic Bus")
    st.caption(f"Logged in as: {st.session_state.get('login_user', '')}")
    page = st.radio("Go to", ["Merge Excel Files", "FAQ / Help"])
    logout_button()

if page == "Merge Excel Files":
    st.title(APP_NAME)
    st.caption("Upload multiple Power BI Excel exports, validate matching headers, remove footer/filter rows, and download one merged Excel file.")

    number_of_files = st.number_input(
        "How many Excel files do you want to upload?",
        min_value=1,
        max_value=100,
        value=10,
        step=1,
    )

    st.info("Please upload Excel files one by one. All files must have the same column headers in the same order.")

    uploaded_files = []
    cols = st.columns(2)
    for i in range(int(number_of_files)):
        with cols[i % 2]:
            file = st.file_uploader(
                f"Upload Excel file {i + 1}",
                type=["xlsx", "xlsm", "xls"],
                key=f"excel_file_{i}",
            )
            uploaded_files.append(file)

    ready_files = [f for f in uploaded_files if f is not None]
    st.write(f"Uploaded files: **{len(ready_files)} / {int(number_of_files)}**")

    if len(ready_files) == int(number_of_files):
        if st.button("Validate, Clean and Merge", type="primary"):
            dataframes = []
            file_summaries = []
            reference_headers = None
            reference_file = None
            header_error = False

            with st.spinner("Reading and validating files..."):
                for idx, file in enumerate(ready_files, start=1):
                    try:
                        df = read_excel_as_text(file)
                        headers = normalize_headers(list(df.columns))

                        if reference_headers is None:
                            reference_headers = headers
                            reference_file = file.name
                        elif headers != reference_headers:
                            header_error = True
                            st.error(f"Header mismatch found in file {idx}: {file.name}")
                            st.write(f"Reference file: **{reference_file}**")
                            comparison = pd.DataFrame(
                                {
                                    "Reference_Header": pd.Series(reference_headers),
                                    "Current_File_Header": pd.Series(headers),
                                }
                            )
                            st.dataframe(comparison, use_container_width=True)
                            break

                        cleaned_df, removed_rows = clean_dataframe(df)
                        dataframes.append(cleaned_df)
                        file_summaries.append(
                            {
                                "File No.": idx,
                                "File Name": file.name,
                                "Original Rows": len(df),
                                "Rows Removed": removed_rows,
                                "Final Rows": len(cleaned_df),
                                "Columns": len(df.columns),
                            }
                        )
                    except Exception as e:
                        header_error = True
                        st.error(f"Could not process file {idx}: {file.name}. Error: {e}")
                        break

            if not header_error:
                merged_df = pd.concat(dataframes, ignore_index=True)
                excel_bytes = build_output_excel(merged_df)
                output_name = f"merged_cleaned_excel_{datetime.now().strftime('%d%m%Y_%H%M%S')}.xlsx"

                st.success("Files validated, cleaned and merged successfully.")
                st.subheader("Processing Summary")
                st.dataframe(pd.DataFrame(file_summaries), use_container_width=True)

                c1, c2, c3 = st.columns(3)
                c1.metric("Files Merged", len(dataframes))
                c2.metric("Final Rows", len(merged_df))
                c3.metric("Columns", len(merged_df.columns))

                st.download_button(
                    label="Download Final Processed Excel",
                    data=excel_bytes,
                    file_name=output_name,
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                )

                with st.expander("Preview merged data"):
                    st.dataframe(merged_df.head(100), use_container_width=True)
    else:
        st.warning("Upload all selected files to start processing.")

else:
    st.title("FAQ / Help")
    st.markdown(
        """
### What does this app do?
This app merges multiple Excel exports into one processed Excel file.

### What is the login?
Default local login:
- Login ID: `admin`
- Password: `MagicBus@2026`

For deployment, change credentials from Streamlit secrets.

### How to change login/password in Streamlit Cloud?
Go to **App settings → Secrets** and add:

```toml
APP_USERNAME = "your_login_id"
APP_PASSWORD = "your_secure_password"
```

### What rows are removed before merging?
Rows are removed when they contain Power BI footer/filter text such as:
- `Total`
- `Applied filters:` and the long filter description below it

### What validation is done?
Before merging, the app checks that all uploaded files have exactly the same headers in the same order. If any file has a different header, the app stops and shows the mismatch.

### Which file formats are supported?
`.xlsx`, `.xlsm`, and `.xls` are accepted. For best results, use `.xlsx`.

### Is data read as text?
Yes. All Excel files are imported using text/string format to avoid date or numeric conversion issues.

### What is the output?
One Excel file with two sheets:
- `Merged_Data`
- `Summary`
        """
    )
