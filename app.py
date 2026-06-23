import io
import hmac
from datetime import datetime
from pathlib import Path
from typing import Dict, List, Tuple

import pandas as pd
import streamlit as st

APP_NAME = "MB Excel Merger"
BASE_DIR = Path(__file__).parent
LOGO_PATH = BASE_DIR / "magicbus_logo.png"

# Region-wise default login for local use.
# For Streamlit Cloud/GitHub deployment, keep the same values in .streamlit/secrets.toml.
DEFAULT_USERS: Dict[str, Dict[str, str]] = {
    "north": {"password": "North@2026", "region": "North"},
    "east": {"password": "East@2026", "region": "East"},
    "west": {"password": "West@2026", "region": "West"},
    "south": {"password": "South@2026", "region": "South"},
}

DROP_TEXTS = {"total", "applied filters:"}


def get_users() -> Dict[str, Dict[str, str]]:
    """Read users from Streamlit secrets. If secrets are missing, use local defaults."""
    try:
        if "users" in st.secrets:
            return {
                str(username): {
                    "password": str(details["password"]),
                    "region": str(details["region"]),
                }
                for username, details in st.secrets["users"].items()
            }
    except Exception:
        pass
    return DEFAULT_USERS


def check_login(username: str, password: str) -> Tuple[bool, str]:
    users = get_users()
    username = username.strip().lower()

    if username not in users:
        return False, ""

    saved_password = users[username]["password"]
    region = users[username]["region"]

    is_valid = hmac.compare_digest(password, saved_password)
    return is_valid, region


def login_screen() -> None:
    st.set_page_config(page_title=f"Login | {APP_NAME}", page_icon="🚌", layout="centered")

    c1, c2, c3 = st.columns([1, 2, 1])
    with c2:
        if LOGO_PATH.exists():
            st.image(str(LOGO_PATH), width=170)
        st.title("Login")
        st.caption(APP_NAME)

        with st.form("login_form"):
            username = st.text_input("Login ID")
            password = st.text_input("Password", type="password")
            submitted = st.form_submit_button("Login", type="primary", use_container_width=True)

        if submitted:
            is_valid, region = check_login(username, password)
            if is_valid:
                st.session_state["authenticated"] = True
                st.session_state["login_user"] = username.strip().lower()
                st.session_state["login_region"] = region
                st.rerun()
            else:
                st.error("Invalid Login ID or Password.")

        with st.expander("Default local users"):
            st.write("North: `north` / `North@2026`")
            st.write("East: `east` / `East@2026`")
            st.write("West: `west` / `West@2026`")
            st.write("South: `south` / `South@2026`")


def logout_button() -> None:
    if st.sidebar.button("Logout"):
        st.session_state.clear()
        st.rerun()


def normalize_headers(headers: List[str]) -> List[str]:
    return [str(h).strip() for h in headers]


def read_excel_as_text(uploaded_file) -> pd.DataFrame:
    return pd.read_excel(uploaded_file, dtype=str, engine="openpyxl")


def row_should_drop(row: pd.Series) -> bool:
    values = ["" if pd.isna(v) else str(v).strip() for v in row.tolist()]
    joined = " ".join(values).strip()
    joined_lower = joined.lower()

    if joined_lower in DROP_TEXTS:
        return True

    if joined_lower.startswith("applied filters:"):
        return True

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


def build_output_excel(df: pd.DataFrame, file_summaries: List[dict], region: str) -> bytes:
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name="Merged_Data")

        summary = pd.DataFrame(
            {
                "Metric": [
                    "Processed On",
                    "Processed By Region",
                    "Total Files Merged",
                    "Total Final Rows",
                    "Total Columns",
                ],
                "Value": [
                    datetime.now().strftime("%d-%m-%Y %H:%M:%S"),
                    region,
                    len(file_summaries),
                    len(df),
                    len(df.columns),
                ],
            }
        )
        summary.to_excel(writer, index=False, sheet_name="Summary")
        pd.DataFrame(file_summaries).to_excel(writer, index=False, sheet_name="File_Wise_Summary")

    output.seek(0)
    return output.getvalue()


if not st.session_state.get("authenticated", False):
    login_screen()
    st.stop()

st.set_page_config(page_title=APP_NAME, page_icon="🚌", layout="wide")

with st.sidebar:
    if LOGO_PATH.exists():
        st.image(str(LOGO_PATH), width=170)
    st.title("Magic Bus")
    st.caption(f"User: {st.session_state.get('login_user', '')}")
    st.caption(f"Region: {st.session_state.get('login_region', '')}")
    page = st.radio("Go to", ["Merge Excel Files", "FAQ / Help"])
    logout_button()

if page == "Merge Excel Files":
    region = st.session_state.get("login_region", "")

    st.title(APP_NAME)
    st.caption("Validate - Clean - Merge - Export")
    st.info(f"You are logged in for **{region}** region.")

    number_of_files = st.number_input(
        "How many Excel files do you want to upload?",
        min_value=1,
        max_value=100,
        value=10,
        step=1,
    )

    st.warning("All uploaded files must have the same headers in the same order.")

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
                                "Region": region,
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
                excel_bytes = build_output_excel(merged_df, file_summaries, region)
                output_name = f"MB_Excel_Merged_{region}_{datetime.now().strftime('%d%m%Y_%H%M%S')}.xlsx"

                st.success("Files validated, cleaned and merged successfully.")
                st.subheader("Processing Summary")
                st.dataframe(pd.DataFrame(file_summaries), use_container_width=True)

                c1, c2, c3, c4 = st.columns(4)
                c1.metric("Region", region)
                c2.metric("Files Merged", len(dataframes))
                c3.metric("Final Rows", len(merged_df))
                c4.metric("Columns", len(merged_df.columns))

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

### Region-wise login
- North: `north` / `North@2026`
- East: `east` / `East@2026`
- West: `west` / `West@2026`
- South: `south` / `South@2026`

### What rows are removed before merging?
Rows are removed when they contain Power BI footer/filter text such as:
- `Total`
- `Applied filters:` and its long filter description

### What validation is done?
Before merging, the app checks that all uploaded files have exactly the same headers in the same order.

### Is data read as text?
Yes. All Excel files are imported as text/string to avoid date or numeric conversion issues.

### Output sheets
The final Excel contains:
- `Merged_Data`
- `Summary`
- `File_Wise_Summary`
        """
    )
