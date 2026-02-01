import tempfile
from pathlib import Path
from datetime import datetime
from zoneinfo import ZoneInfo

import streamlit as st

from transform import transform_files


# --- Config ---
TEMPLATE_PATH = Path("templates") / "Renee(B).xlsx"
NY_TZ = ZoneInfo("America/New_York")


st.set_page_config(page_title="Renee Supplier Data Transformer", layout="wide")

st.title("Renee Supplier Data Transformer")
st.write("Upload the supplier sheet (Spreadsheet A). The app outputs an Excel file in the fixed **Renee(B)** format.")


with st.sidebar:
    st.header("Options")
    defect_exclusion = st.checkbox(
        "Exclude quantities that contain Chinese/non-ASCII annotations (e.g., 瑕疵)",
        value=True
    )
    filldown_b_template = st.checkbox(
        "Fill down missing Model/Blade cells in B template (recommended)",
        value=True
    )
    make_diff = st.checkbox("Generate diff CSV vs template B", value=False)

    st.divider()
    st.caption("Advanced (optional)")
    override_template = st.file_uploader(
        "Optional: Upload a different B template (.xlsx)",
        type=["xlsx"],
        help="Leave blank to use the repo template at templates/Renee(B).xlsx"
    )


# --- Validate template existence ---
if not TEMPLATE_PATH.exists() and override_template is None:
    st.error(
        f"Missing template file: {TEMPLATE_PATH}\n\n"
        "Fix: Put Renee(B).xlsx at templates/Renee(B).xlsx or upload a template in the sidebar."
    )
    st.stop()


st.divider()

file_a = st.file_uploader(
    "Upload Spreadsheet A (.xlsx) (can be named anything)",
    type=["xlsx"],
    key="a"
)

run = st.button("Transform", type="primary", disabled=(file_a is None))

if run:
    with st.spinner("Transforming..."):
        with tempfile.TemporaryDirectory() as tmpdir:
            tmpdir = Path(tmpdir)

            # Save A
            path_a = tmpdir / "input_A.xlsx"
            path_a.write_bytes(file_a.getbuffer())

            # Choose template B: override upload or repo template
            if override_template is not None:
                path_b = tmpdir / "template_B.xlsx"
                path_b.write_bytes(override_template.getbuffer())
            else:
                path_b = TEMPLATE_PATH

            # Output name: Stick_List_Date (today in America/New_York)
            today_str = datetime.now(NY_TZ).date().strftime("%Y-%m-%d")
            out_filename = f"Stick_List_{today_str}.xlsx"
            out_xlsx = tmpdir / out_filename

            diff_csv = tmpdir / "diff_report.csv" if make_diff else None

            transform_files(
                path_a=str(path_a),
                path_b=str(path_b),
                out_path=str(out_xlsx),
                defect_exclusion=defect_exclusion,
                filldown_b_template=filldown_b_template,
                diff_csv=(str(diff_csv) if diff_csv else None),
            )

            st.success("Done!")

            st.download_button(
                label=f"Download {out_filename}",
                data=out_xlsx.read_bytes(),
                file_name=out_filename,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )

            if make_diff and diff_csv and diff_csv.exists():
                st.download_button(
                    label="Download diff_report.csv",
                    data=diff_csv.read_bytes(),
                    file_name="diff_report.csv",
                    mime="text/csv",
                )