import base64
import hmac
import io
from datetime import datetime
from pathlib import Path
from typing import Dict, List, Tuple

import pandas as pd
import streamlit as st

APP_NAME = "MB Excel Merger"
APP_TAGLINE = "Validate • Clean • Merge • Export"
BASE_DIR = Path(__file__).parent
LOGO_PATH = BASE_DIR / "magicbus_logo.png"
MAGIC_BUS_URL = "https://www.magicbus.org/"

# Login is configured from .streamlit/secrets.toml or Streamlit Cloud Secrets.
# For private/internal GitHub repositories, update values in Streamlit Cloud Secrets.
DEFAULT_USERS: Dict[str, Dict[str, str]] = {
    "north": {"password": "North@2026", "region": "North"},
    "east": {"password": "East@2026", "region": "East"},
    "west": {"password": "West@2026", "region": "West"},
    "south": {"password": "South@2026", "region": "South"},
}

DROP_TEXTS = {"total", "applied filters:"}
SUPPORTED_TYPES = ["csv", "xls"]


def inject_css() -> None:
    st.markdown(
        """
        <style>
            .main .block-container {padding-top: 1.2rem; padding-bottom: 2rem; max-width: 1180px;}
            .mb-hero {
                background: linear-gradient(135deg, #fff7e6 0%, #ffffff 58%, #fff1cc 100%);
                border: 1px solid #f2d28b;
                border-radius: 22px;
                padding: 22px 24px;
                margin-bottom: 18px;
                box-shadow: 0 8px 28px rgba(0,0,0,0.06);
            }
            .mb-logo-card {
                background: #ffffff;
                border: 1px solid #eee3c4;
                border-radius: 18px;
                padding: 14px;
                text-align: center;
                box-shadow: 0 4px 16px rgba(0,0,0,0.05);
            }
            .mb-logo-card img {max-width: 150px; height: auto;}
            .mb-logo-card small {display: block; color: #6b6b6b; margin-top: 6px;}
            .mb-title {font-size: 2.15rem; font-weight: 800; margin: 0; color: #1f2937;}
            .mb-subtitle {font-size: 1.02rem; color: #5f6673; margin-top: 4px;}
            .mb-card {
                background: #ffffff;
                border: 1px solid #e9edf2;
                border-radius: 18px;
                padding: 18px;
                box-shadow: 0 4px 18px rgba(0,0,0,0.04);
                min-height: 122px;
            }
            .mb-card h4 {margin: 0 0 8px 0; color: #1f2937;}
            .mb-card p {margin: 0; color: #606975; line-height: 1.45;}
            .mb-pill {
                display: inline-block; padding: 7px 12px; border-radius: 999px;
                background: #fff3cd; color: #5c4500; font-weight: 700; font-size: 0.88rem;
                border: 1px solid #f5d987; margin-right: 8px; margin-bottom: 8px;
            }
            .footer-note {color:#6b7280; font-size:0.9rem; margin-top:24px;}
        </style>
        """,
        unsafe_allow_html=True,
    )


def logo_html(width: int = 150) -> str:
    if not LOGO_PATH.exists():
        return ""
    encoded = base64.b64encode(LOGO_PATH.read_bytes()).decode("utf-8")
    return f"""
    <div class="mb-logo-card">
        <a href="{MAGIC_BUS_URL}" target="_blank" title="Open Magic Bus website">
            <img src="data:image/png;base64,{encoded}" width="{width}" />
        </a>
        <small>Click logo to visit Magic Bus</small>
    </div>
    """


def render_header(region: str | None = None) -> None:
    left, right = st.columns([1.15, 5])
    with left:
        st.markdown(logo_html(), unsafe_allow_html=True)
    with right:
        st.markdown(
            f"""
            <div class="mb-hero">
                <p class="mb-title">{APP_NAME}</p>
                <p class="mb-subtitle">{APP_TAGLINE}</p>
                <div style="margin-top: 10px;">
                    <span class="mb-pill">Same header validation</span>
                    <span class="mb-pill">Power BI footer cleanup</span>
                    <span class="mb-pill">CSV / XLS only</span>
                    {f'<span class="mb-pill">Region: {region}</span>' if region else ''}
                </div>
            </div>
            """,
            unsafe_allow_html=True,
        )


def get_users() -> Dict[str, Dict[str, str]]:
    try:
        if "users" in st.secrets:
            return {
                str(username).strip().lower(): {
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
    return hmac.compare_digest(password, saved_password), region


def login_screen() -> None:
    st.set_page_config(page_title=f"Login | {APP_NAME}", page_icon="🚌", layout="centered")
    inject_css()
    render_header()

    c1, c2, c3 = st.columns([1, 1.4, 1])
    with c2:
        st.subheader("Secure Login")
        st.caption("Use your region-wise login ID to access the merger utility.")
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


def logout_button() -> None:
    if st.sidebar.button("Logout", use_container_width=True):
        st.session_state.clear()
        st.rerun()


def normalize_headers(headers: List[str]) -> List[str]:
    return [str(h).strip() for h in headers]


def read_file_as_text(uploaded_file) -> pd.DataFrame:
    file_name = uploaded_file.name.lower()
    if file_name.endswith(".csv"):
        return pd.read_csv(uploaded_file, dtype=str, keep_default_na=False)
    if file_name.endswith(".xls"):
        return pd.read_excel(uploaded_file, dtype=str, engine="xlrd", keep_default_na=False)
    raise ValueError("Unsupported file type. Please upload only .csv or .xls files.")


def row_should_drop(row: pd.Series) -> bool:
    values = ["" if pd.isna(v) else str(v).strip() for v in row.tolist()]
    joined_lower = " ".join(values).strip().lower()
    if joined_lower in DROP_TEXTS:
        return True
    if joined_lower.startswith("applied filters:"):
        return True
    non_blank = [v for v in values if v != ""]
    if len(non_blank) == 1:
        only_value = non_blank[0].strip().lower()
        return only_value == "total" or only_value.startswith("applied filters:")
    return False


def clean_dataframe(df: pd.DataFrame) -> Tuple[pd.DataFrame, int]:
    before = len(df)
    df = df.dropna(how="all")
    mask = df.apply(row_should_drop, axis=1)
    df = df.loc[~mask].copy()
    return df, before - len(df)


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
                    "Input File Types Allowed",
                ],
                "Value": [
                    datetime.now().strftime("%d-%m-%Y %H:%M:%S"),
                    region,
                    len(file_summaries),
                    len(df),
                    len(df.columns),
                    ".csv, .xls",
                ],
            }
        )
        summary.to_excel(writer, index=False, sheet_name="Summary")
        pd.DataFrame(file_summaries).to_excel(writer, index=False, sheet_name="File_Wise_Summary")
    output.seek(0)
    return output.getvalue()


def render_why_page() -> None:
    st.subheader("Why this app is being used")
    c1, c2, c3 = st.columns(3)
    with c1:
        st.markdown("""<div class="mb-card"><h4>1. Standard process</h4><p>Every region follows the same controlled workflow for combining files.</p></div>""", unsafe_allow_html=True)
    with c2:
        st.markdown("""<div class="mb-card"><h4>2. Fewer manual errors</h4><p>The app checks headers before merging, so wrong-format files are stopped early.</p></div>""", unsafe_allow_html=True)
    with c3:
        st.markdown("""<div class="mb-card"><h4>3. Cleaner output</h4><p>Power BI footer rows like Total and Applied filters are removed before append.</p></div>""", unsafe_allow_html=True)

    st.markdown("### How it saves time")
    st.markdown(
        """
        Instead of opening every file manually, deleting filter rows, checking columns, copying data, and appending it in Excel, users can upload all files in one place. The app validates the structure, cleans unwanted rows, merges the data, and creates a ready-to-use processed Excel output with a summary sheet.
        """
    )

    st.markdown("### Important file rule")
    st.warning("This app currently accepts only `.csv` and `.xls` files. Other formats will not be processed.")


def render_faq_page() -> None:
    st.subheader("FAQ / Help")
    with st.expander("What does this app do?", expanded=True):
        st.write("It validates, cleans, and merges multiple region-wise data files into one processed Excel output.")
    with st.expander("Why was this app created?"):
        st.write("It reduces repetitive manual Excel work, prevents header mismatch errors, and creates a standard merge process for all regions.")
    with st.expander("Which file formats are supported?"):
        st.write("Only `.csv` and `.xls` files are supported in the upload section.")
    with st.expander("What validation is done before merging?"):
        st.write("All uploaded files must have exactly the same headers in the same order. If any file has different headers, processing stops and a comparison table is shown.")
    with st.expander("What rows are removed before merging?"):
        st.write("Rows containing `Total` and rows starting with `Applied filters:` are removed before data is appended.")
    with st.expander("What does the final output contain?"):
        st.write("The final Excel contains three sheets: `Merged_Data`, `Summary`, and `File_Wise_Summary`.")
    with st.expander("Will it change date or number values?"):
        st.write("The app reads all columns as text to reduce unwanted date/number conversion issues during merging.")


def render_merge_page() -> None:
    region = st.session_state.get("login_region", "")
    st.info(f"You are logged in for **{region}** region.")

    c1, c2, c3, c4 = st.columns(4)
    c1.metric("Allowed formats", ".csv / .xls")
    c2.metric("Validation", "Header match")
    c3.metric("Cleanup", "Total / Filters")
    c4.metric("Output", ".xlsx")

    with st.container(border=True):
        st.subheader("Upload files")
        number_of_files = st.number_input(
            "How many files do you want to upload?",
            min_value=1,
            max_value=100,
            value=10,
            step=1,
        )
        st.caption("Upload the exact number of files selected above. Processing starts only after all file pickers have files.")
        st.warning("All uploaded files must have the same headers in the same order.")

        uploaded_files = []
        cols = st.columns(2)
        for i in range(int(number_of_files)):
            with cols[i % 2]:
                file = st.file_uploader(
                    f"Upload file {i + 1}",
                    type=SUPPORTED_TYPES,
                    key=f"data_file_{i}",
                )
                uploaded_files.append(file)

    ready_files = [f for f in uploaded_files if f is not None]
    st.progress(len(ready_files) / int(number_of_files), text=f"Uploaded files: {len(ready_files)} / {int(number_of_files)}")

    if len(ready_files) != int(number_of_files):
        st.warning("Upload all selected files to start processing.")
        return

    if not st.button("Validate, Clean and Merge", type="primary", use_container_width=True):
        return

    dataframes = []
    file_summaries = []
    reference_headers = None
    reference_file = None
    header_error = False

    with st.spinner("Reading, validating, cleaning, and merging files..."):
        for idx, file in enumerate(ready_files, start=1):
            try:
                df = read_file_as_text(file)
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

    if header_error:
        return

    merged_df = pd.concat(dataframes, ignore_index=True)
    excel_bytes = build_output_excel(merged_df, file_summaries, region)
    output_name = f"MB_Excel_Merged_{region}_{datetime.now().strftime('%d%m%Y_%H%M%S')}.xlsx"

    st.success("Files validated, cleaned and merged successfully.")
    st.subheader("Processing Summary")
    st.dataframe(pd.DataFrame(file_summaries), use_container_width=True)

    c1, c2, c3, c4 = st.columns(4)
    c1.metric("Region", region)
    c2.metric("Files Merged", len(dataframes))
    c3.metric("Final Rows", f"{len(merged_df):,}")
    c4.metric("Columns", len(merged_df.columns))

    st.download_button(
        label="Download Final Processed Excel",
        data=excel_bytes,
        file_name=output_name,
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True,
    )

    with st.expander("Preview merged data"):
        st.dataframe(merged_df.head(100), use_container_width=True)


if not st.session_state.get("authenticated", False):
    login_screen()
    st.stop()

st.set_page_config(page_title=APP_NAME, page_icon="🚌", layout="wide")
inject_css()

with st.sidebar:
    st.markdown(logo_html(width=130), unsafe_allow_html=True)
    st.markdown("### MB Excel Merger")
    st.caption(f"User: {st.session_state.get('login_user', '')}")
    st.caption(f"Region: {st.session_state.get('login_region', '')}")
    page = st.radio("Navigation", ["Merge Files", "Why this app", "FAQ / Help"], label_visibility="collapsed")
    logout_button()

render_header(region=st.session_state.get("login_region", ""))

if page == "Merge Files":
    render_merge_page()
elif page == "Why this app":
    render_why_page()
else:
    render_faq_page()

st.markdown("<p class='footer-note'>Internal utility for standard data merge workflow. For support, contact the data systems team.</p>", unsafe_allow_html=True)
