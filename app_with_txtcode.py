import streamlit as st
import subprocess
from pathlib import Path
import zipfile
import re
import shutil
import os

st.set_page_config(page_title="Design Export Automation", layout="wide")

st.title("Design Export Automation")
st.write("Enter inputs, choose folder path or upload PDFs, run automation, and download the generated Excel files.")

# ================================================================
# STATIC PATHS - CHANGE HERE ONLY IF NEEDED
# ================================================================
EXCEL_FILE = r"template/DESIGN EXPORT TEMPLATE.xlsx"
DUCKDB_FILE = r"data/uk_risk.duckdb"
OUTPUT_ROOT = r"output_folders"

# Temporary folder used only when PDFs are uploaded manually
UPLOAD_INPUT_ROOT = r"uploaded_input_folders"

# ================================================================
# ORIGINAL SCRIPT - DO NOT EDIT PROCESSING LOGIC
# ================================================================
with open("original_script.txt", "r", encoding="utf-8") as f:
    ORIGINAL_CODE = f.read()


# ================================================================
# STREAMLIT INPUTS
# ================================================================
st.sidebar.header("Inputs")

SS_TOKEN = st.sidebar.text_input("SS_TOKEN", type="password")
SHEET_ID = st.sidebar.text_input("SHEET_ID")
SECOND_SHEET_ID = st.sidebar.text_input("SECOND_SHEET_ID")

static_local_authority = st.sidebar.text_input("static_local_authority")

SHEET10_B3_VALUE = st.sidebar.text_input("SHEET10_B3_VALUE")
SHEET10_F3_VALUE = st.sidebar.text_input("SHEET10_F3_VALUE")

st.sidebar.divider()

input_type = st.sidebar.radio(
    "Choose PDF input type",
    ["Folder path", "Upload PDFs manually"]
)

PARENT_FOLDER = ""

if input_type == "Folder path":
    PARENT_FOLDER = st.sidebar.text_input(
        "PARENT_FOLDER",
        help="Paste the folder path containing property folders. Each property folder should contain a Retrofit folder."
    )

else:
    property_name = st.sidebar.text_input(
        "Property folder name",
        value="Uploaded_Property",
        help="This name will be used for the output Excel file names."
    )

    uploaded_pdfs = st.file_uploader(
        "Upload CR, SN and EPR PDF files",
        type=["pdf"],
        accept_multiple_files=True
    )


# ================================================================
# HELPERS
# ================================================================
def py_string(value):
    return repr(str(value))


def py_int(value, field_name):
    value = str(value).strip()
    if not value.isdigit():
        raise ValueError(f"{field_name} must be a number")
    return value


def replace_assignment(code, name, value_code):
    pattern = rf"^(\s*){re.escape(name)}\s*=.*$"
    replacement = rf"\1{name} = {value_code}"

    code, count = re.subn(
        pattern,
        replacement,
        code,
        count=1,
        flags=re.MULTILINE
    )

    # If variable is not found in original_script.txt, add it at the top
    if count == 0:
        code = f"{name} = {value_code}\n" + code

    return code


def clear_folder(folder_path):
    folder = Path(folder_path)
    if folder.exists():
        shutil.rmtree(folder)
    folder.mkdir(parents=True, exist_ok=True)


def prepare_uploaded_pdf_folder(uploaded_files, property_name):
    if not uploaded_files:
        raise ValueError("Please upload CR, SN and EPR PDF files.")

    if len(uploaded_files) < 3:
        raise ValueError("Please upload at least 3 PDFs: CR, SN and EPR.")

    safe_property_name = re.sub(r"[^A-Za-z0-9 _-]", "", property_name).strip()
    if not safe_property_name:
        safe_property_name = "Uploaded_Property"

    input_root = Path(UPLOAD_INPUT_ROOT)
    input_root.mkdir(parents=True, exist_ok=True)

    run_folder = input_root / f"run_{os.getpid()}"
    property_folder = run_folder / safe_property_name
    retrofit_folder = property_folder / "Retrofit"
    retrofit_folder.mkdir(parents=True, exist_ok=True)

    for uploaded_file in uploaded_files:
        file_path = retrofit_folder / uploaded_file.name

        with open(file_path, "wb") as f:
            f.write(uploaded_file.getbuffer())

    return str(run_folder.resolve())

def build_runtime_script(parent_folder_value):
    code = ORIGINAL_CODE

    # Only replacing requested inputs
    code = replace_assignment(code, "SS_TOKEN", py_string(SS_TOKEN))
    code = replace_assignment(code, "SHEET_ID", py_int(SHEET_ID, "SHEET_ID"))
    code = replace_assignment(code, "SECOND_SHEET_ID", py_int(SECOND_SHEET_ID, "SECOND_SHEET_ID"))
    code = replace_assignment(code, "static_local_authority", py_string(static_local_authority))
    code = replace_assignment(code, "SHEET10_B3_VALUE", py_string(SHEET10_B3_VALUE))
    code = replace_assignment(code, "SHEET10_F3_VALUE", py_string(SHEET10_F3_VALUE))

    # PARENT_FOLDER can come from folder path OR uploaded PDFs temp folder
    code = replace_assignment(code, "PARENT_FOLDER", py_string(parent_folder_value))

    # Static paths from this Streamlit file
    code = replace_assignment(code, "EXCEL_FILE", py_string(EXCEL_FILE))
    code = replace_assignment(code, "DUCKDB_FILE", py_string(DUCKDB_FILE))
    code = replace_assignment(code, "OUTPUT_ROOT", py_string(OUTPUT_ROOT))

    return code


def zip_excel_files(files, output_root):
    zip_path = Path(output_root) / "Design_Export_Outputs.zip"

    if zip_path.exists():
        zip_path.unlink()

    with zipfile.ZipFile(zip_path, "w", zipfile.ZIP_DEFLATED) as zf:
        for file in files:
            zf.write(file, arcname=file.name)

    return zip_path


def clean_previous_outputs():
    output_path = Path(OUTPUT_ROOT)
    output_path.mkdir(parents=True, exist_ok=True)

    for file in output_path.glob("*.xlsx"):
        file.unlink()

    old_zip = output_path / "Design_Export_Outputs.zip"
    if old_zip.exists():
        old_zip.unlink()


# ================================================================
# RUN BUTTON
# ================================================================
if st.button("Run Automation", type="primary"):

    missing = []

    if not SS_TOKEN.strip():
        missing.append("SS_TOKEN")
    if not SHEET_ID.strip():
        missing.append("SHEET_ID")
    if not SECOND_SHEET_ID.strip():
        missing.append("SECOND_SHEET_ID")

    if input_type == "Folder path":
        if not PARENT_FOLDER.strip():
            missing.append("PARENT_FOLDER")
    else:
        if not property_name.strip():
            missing.append("Property folder name")
        if not uploaded_pdfs:
            missing.append("CR/SN/EPR PDFs")

    if missing:
        st.error("Please fill: " + ", ".join(missing))
        st.stop()

    Path(OUTPUT_ROOT).mkdir(parents=True, exist_ok=True)

    runtime_script = Path(OUTPUT_ROOT) / "_runtime_design_export.py"

    with st.status("Running automation...", expanded=True) as status:
        try:
            clean_previous_outputs()

            if input_type == "Upload PDFs manually":
                parent_folder_to_use = prepare_uploaded_pdf_folder(uploaded_pdfs, property_name)
                st.info(f"Uploaded PDFs saved into temporary input folder: {parent_folder_to_use}")
            else:
                parent_folder_to_use = PARENT_FOLDER

            runtime_code = build_runtime_script(parent_folder_to_use)
            runtime_script.write_text(runtime_code, encoding="utf-8")

            result = subprocess.run(
                ["python", str(runtime_script)],
                capture_output=True,
                text=True,
                encoding="utf-8",
                errors="replace",
                env={
                    **os.environ,
                    "PYTHONIOENCODING": "utf-8"
                }
            )

            if result.stdout:
                st.subheader("Run Log")
                st.code(result.stdout)

            if result.stderr:
                st.subheader("Errors / Warnings")
                st.code(result.stderr)

            if result.returncode != 0:
                status.update(label="Automation failed", state="error")
                st.error("Script failed. Check error above.")
                st.stop()

            excel_files = sorted(
                [
                    file for file in Path(OUTPUT_ROOT).glob("*.xlsx")
                    if file.name.endswith("_Prelim.xlsx") or file.name.endswith("_Construction.xlsx")
                ],
                key=lambda x: x.name
            )

            if not excel_files:
                status.update(label="Completed, but no Excel files found", state="error")
                st.warning("No Prelim or Construction Excel files found.")
                st.stop()

            status.update(label="Automation completed", state="complete")
            st.success(f"Generated {len(excel_files)} Excel file(s).")

            st.subheader("Download Excel Files")

            for file in excel_files:
                with open(file, "rb") as f:
                    st.download_button(
                        label=f"Download {file.name}",
                        data=f,
                        file_name=file.name,
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    )

            zip_path = zip_excel_files(excel_files, OUTPUT_ROOT)

            with open(zip_path, "rb") as f:
                st.download_button(
                    label="Download All Excel Files as ZIP",
                    data=f,
                    file_name=zip_path.name,
                    mime="application/zip"
                )

        except Exception as e:
            status.update(label="Automation failed", state="error")
            st.exception(e)
