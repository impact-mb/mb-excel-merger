import base64
import hmac
import io
from datetime import datetime
from pathlib import Path

import pandas as pd
import streamlit as st


APP_NAME = "MB Excel Merger"
BASE_DIR = Path(__file__).parent
LOGO_PATH = BASE_DIR / "magicbus_logo.png"
MAGIC_BUS_URL = "https://www.magicbus.org/"

SUPPORTED_TYPES = ["csv", "xls", "xlsx"]
DROP_TEXTS = {"total", "applied filters:"}


st.set_page_config(
    page_title=APP_NAME,
    page_icon="🚌",
    layout="wide"
)


def get_users():
    try:
        users = {}
        for username, details in st.secrets["users"].items():
            users[str(username).strip().lower()] = {
                "password": str(details["password"]),
                "region": str(details["region"])
            }
        return users
    except Exception:
        st.error("Login users are not configured in Streamlit Secrets.")
        return {}


def check_login(username, password):
    users = get_users()
    username = username.strip().lower()

    if username not in users:
        return False, ""

    stored_password = users[username]["password"]
    region = users[username]["region"]

    return hmac.compare_digest(password, stored_password), region


def logo_html(width=150):
    if not LOGO_PATH.exists():
        return ""

    encoded_logo = base64.b64encode(LOGO_PATH.read_bytes()).decode("utf-8")

    return f"""
    <a href="{MAGIC_BUS_URL}" target="_blank">
        <img src="data:image/png;base64,{encoded_logo}" width="{width}">
    </a>
    """


def login_page():
    st.markdown(logo_html(170), unsafe_allow_html=True)
    st.title("MB Excel Merger")
    st.caption("Login to merge and export your files.")

    with st.form("login_form"):
        username = st.text_input("Login ID")
        password = st.text_input("Password", type="password")
        login = st.form_submit_button("Login", type="primary")

    if login:
        valid, region = check_login(username, password)

        if valid:
            st.session_state["authenticated"] = True
            st.session_state["login_user"] = username.strip().lower()
            st.session_state["login_region"] = region
            st.rerun()
        else:
            st.error("Invalid Login ID or Password.")


def read_file(uploaded_file):

    file_name = uploaded_file.name.lower()

    if file_name.endswith(".csv"):

        return pd.read_csv(
            uploaded_file,
            dtype=str,
            keep_default_na=False
        )

    elif file_name.endswith(".xls"):

        return pd.read_excel(
            uploaded_file,
            dtype=str,
            engine="xlrd",
            keep_default_na=False
        )

    elif file_name.endswith(".xlsx"):

        return pd.read_excel(
            uploaded_file,
            dtype=str,
            engine="openpyxl",
            keep_default_na=False
        )

    else:

        raise ValueError(
            "Only .csv, .xls and .xlsx files are supported."
        )


def normalize_headers(headers):
    return [str(h).strip() for h in headers]


def row_should_drop(row):
    values = ["" if pd.isna(v) else str(v).strip() for v in row.tolist()]
    joined_text = " ".join(values).strip().lower()

    if joined_text in DROP_TEXTS:
        return True

    if joined_text.startswith("applied filters:"):
        return True

    non_blank = [v for v in values if v != ""]

    if len(non_blank) == 1:
        only_value = non_blank[0].strip().lower()

        if only_value == "total":
            return True

        if only_value.startswith("applied filters:"):
            return True

    return False


def clean_dataframe(df):
    original_rows = len(df)

    df = df.dropna(how="all")

    drop_mask = df.apply(row_should_drop, axis=1)

    cleaned_df = df.loc[~drop_mask].copy()

    removed_rows = original_rows - len(cleaned_df)

    return cleaned_df, removed_rows


def build_excel_output(merged_df, summary_df, region):
    output = io.BytesIO()

    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        merged_df.to_excel(writer, index=False, sheet_name="Merged_Data")
        summary_df.to_excel(writer, index=False, sheet_name="File_Summary")

        app_summary = pd.DataFrame({
            "Metric": [
                "Processed Date",
                "Processed Region",
                "Total Files Merged",
                "Total Final Rows",
                "Total Columns",
                "Allowed File Types"
            ],
            "Value": [
                datetime.now().strftime("%d-%m-%Y"),
                region,
                len(summary_df),
                len(merged_df),
                len(merged_df.columns),
                ".csv, .xls"
            ]
        })

        app_summary.to_excel(writer, index=False, sheet_name="Summary")

    output.seek(0)
    return output.getvalue()


def main_app():
    region = st.session_state.get("login_region", "")

    col1, col2 = st.columns([1, 5])

    with col1:
        st.markdown(logo_html(140), unsafe_allow_html=True)

    with col2:
        st.title("MB Excel Merger")
        st.caption("Upload files → Validate headers → Clean unwanted rows → Export final Excel")

    st.divider()

    st.info(f"Logged in Region: **{region}**")

    if st.button("Logout"):
        st.session_state.clear()
        st.rerun()

    st.subheader("Upload Files")

    number_of_files = st.number_input(
        "How many files do you want to upload?",
        min_value=1,
        max_value=100,
        value=10,
        step=1
    )

    st.warning("Only `.csv` and `.xls` files are supported. All files must have the same headers.")

    uploaded_files = []

    for i in range(int(number_of_files)):
        file = st.file_uploader(
            f"Upload file {i + 1}",
            type=SUPPORTED_TYPES,
            key=f"file_{i}"
        )
        uploaded_files.append(file)

    ready_files = [file for file in uploaded_files if file is not None]

    st.progress(
        len(ready_files) / int(number_of_files),
        text=f"Uploaded {len(ready_files)} of {int(number_of_files)} files"
    )

    if len(ready_files) != int(number_of_files):
        st.stop()

    if not st.button("Validate, Clean and Merge", type="primary"):
        st.stop()

    dataframes = []
    summary_list = []

    reference_headers = None
    reference_file = None

    for index, file in enumerate(ready_files, start=1):
        try:
            df = read_file(file)

            current_headers = normalize_headers(df.columns.tolist())

            if reference_headers is None:
                reference_headers = current_headers
                reference_file = file.name

            elif current_headers != reference_headers:
                st.error(f"Header mismatch found in file: {file.name}")
                st.write(f"Reference file: **{reference_file}**")

                comparison = pd.DataFrame({
                    "Reference Header": pd.Series(reference_headers),
                    "Current File Header": pd.Series(current_headers)
                })

                st.dataframe(comparison, use_container_width=True)
                st.stop()

            cleaned_df, removed_rows = clean_dataframe(df)

            dataframes.append(cleaned_df)

            summary_list.append({
                "File No.": index,
                "File Name": file.name,
                "Original Rows": len(df),
                "Rows Removed": removed_rows,
                "Final Rows": len(cleaned_df),
                "Columns": len(df.columns)
            })

        except Exception as e:
            st.error(f"Error in file {file.name}: {e}")
            st.stop()

    merged_df = pd.concat(dataframes, ignore_index=True)

    summary_df = pd.DataFrame(summary_list)

    excel_file = build_excel_output(merged_df, summary_df, region)

    current_date = datetime.now().strftime("%d-%m-%Y")

    output_file_name = (f"MB_Excel_Merged_{region}_{current_date}.xlsx")
    
    st.success("Files merged successfully.")

    st.subheader("Processing Summary")
    st.dataframe(summary_df, use_container_width=True)

    c1, c2, c3 = st.columns(3)

    c1.metric("Files Merged", len(dataframes))
    c2.metric("Final Rows", f"{len(merged_df):,}")
    c3.metric("Final Columns", len(merged_df.columns))

    st.download_button(
        label="Download Final Excel",
        data=excel_file,
        file_name=output_file_name,
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        type="primary"
    )


if not st.session_state.get("authenticated", False):
    login_page()
else:
    main_app()