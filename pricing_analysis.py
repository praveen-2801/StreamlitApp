import streamlit as st
from openpyxl import load_workbook, Workbook
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
from io import BytesIO
from copy import copy

st.set_page_config(page_title="Pricing Tool", layout="wide")

st.sidebar.title("Navigation")
selected_option = st.sidebar.radio("Choose an option", ["Home", "Pricing Analysis Version - 2"])

def upload_files():
    st.write("Upload Files")
    uploaded = st.file_uploader("Upload Excel files", type=["xlsx"], accept_multiple_files=True)
    if uploaded:
        st.success(f"{len(uploaded)} file(s) uploaded successfully.")
    return uploaded

# Extract data from each Excel file using Reference Number
def extract_data_by_reference(file):
    in_memory_file = BytesIO(file.read())
    wb = load_workbook(in_memory_file, data_only=False)
    ws = wb.active
    rows = list(ws.iter_rows(values_only=False))

    headers = [cell.value for cell in rows[0]]
    ref_col_idx = next((idx for idx, val in enumerate(headers) if val and "reference" in str(val).lower()), None)
    if ref_col_idx is None:
        raise ValueError("Reference Number column not found")

    header_cells = []
    for cell in rows[0]:
        cell_data = {
            "value": cell.value,
            "fill": copy(cell.fill),
            "font": copy(cell.font),
            "alignment": copy(cell.alignment),
            "border": copy(cell.border)
        }
        header_cells.append(cell_data)

    formatted_data = {}
    for row in rows[1:]:
        ref_val = row[ref_col_idx].value
        if not ref_val:
            continue
        cell_data_row = []
        for cell in row:
            cell_data_row.append({
                "value": cell.value,
                "fill": copy(cell.fill),
                "font": copy(cell.font),
                "alignment": copy(cell.alignment),
                "border": copy(cell.border)
            })
        formatted_data[ref_val] = cell_data_row

    return header_cells, ref_col_idx, formatted_data

# Merge data from all files
def merge_files_by_reference(uploaded_files):
    all_references = set()
    file_data = []

    for file in uploaded_files:
        headers, ref_idx, data = extract_data_by_reference(file)
        file_data.append((headers, ref_idx, data))
        all_references.update(data.keys())

    all_references = sorted(all_references)
    return all_references, file_data

# Create bordered cell
def apply_border(cell):
    thin_border = Border(
        left=Side(style='thin'),
        right=Side(style='thin'),
        top=Side(style='thin'),
        bottom=Side(style='thin')
    )
    cell.border = thin_border

# Create final workbook
def create_workbook(all_refs, file_data):
    wb = Workbook()
    ws = wb.active
    ws.title = "Merged"

    final_headers = ["Reference Number"]
    header_styles = []

    for headers, ref_idx, _ in file_data:
        for idx, cell in enumerate(headers):
            if idx != ref_idx:
                final_headers.append(cell["value"])
                header_styles.append(cell)

    # Add new columns
    final_headers += ["Annual Volume/100", "x"]

    # Write headers
    ref_header_style = file_data[0][0][file_data[0][1]]
    cell = ws.cell(row=1, column=1, value="Reference Number")
    cell.fill = ref_header_style["fill"]
    cell.font = ref_header_style["font"]
    cell.alignment = ref_header_style["alignment"]
    apply_border(cell)

    for i, h in enumerate(header_styles):
        cell = ws.cell(row=1, column=i + 2, value=h["value"])
        cell.fill = h["fill"]
        cell.font = h["font"]
        cell.alignment = h["alignment"]
        apply_border(cell)

    # Add headers for new columns
    new_headers = ["Annual Volume/100", "x"]
    for i, new_header in enumerate(new_headers, start=len(header_styles) + 2):
        cell = ws.cell(row=1, column=i, value=new_header)
        cell.font = Font(bold=True)
        cell.alignment = Alignment(horizontal="center", vertical="center")
        apply_border(cell)

    # Write data
    for row_idx, ref_val in enumerate(all_refs, start=2):
        cell = ws.cell(row=row_idx, column=1, value=ref_val)
        apply_border(cell)

        col_idx = 2
        for headers, ref_idx, data in file_data:
            if ref_val in data:
                row_data = data[ref_val]
                for i, cell_obj in enumerate(row_data):
                    if i == ref_idx:
                        continue
                    value = cell_obj["value"] if cell_obj["value"] is not None else None
                    cell = ws.cell(row=row_idx, column=col_idx, value=value)
                    cell.fill = cell_obj["fill"]
                    cell.font = cell_obj["font"]
                    cell.alignment = cell_obj["alignment"]
                    apply_border(cell)
                    col_idx += 1
            else:
                for i in range(len(headers)):
                    if i != ref_idx:
                        cell = ws.cell(row=row_idx, column=col_idx, value=None)
                        apply_border(cell)
                        col_idx += 1

        # Add placeholders for new columns
        ws.cell(row=row_idx, column=col_idx, value=None)
        apply_border(ws.cell(row=row_idx, column=col_idx))
        ws.cell(row=row_idx, column=col_idx + 1, value=None)
        apply_border(ws.cell(row=row_idx, column=col_idx + 1))

    return wb


# Download link
def get_download_link_with_styles(wb):
    output = BytesIO()
    wb.save(output)
    output.seek(0)
    st.download_button(
        label="Download Merged Excel File by Reference",
        data=output,
        file_name="reference_merged_output.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

# UI Logic
def pricing_analysis_v2():
    st.title("Pricing Analysis Version - 2 (Reference-Based Merge with Formatting)")

    uploaded_files = upload_files()
    if uploaded_files and st.button("Merge Files"):
        try:
            all_refs, file_data = merge_files_by_reference(uploaded_files)
            wb = create_workbook(all_refs, file_data)
            st.success("Files merged successfully with formatting and borders.")
            get_download_link_with_styles(wb)
        except Exception as e:
            st.error(f"An error occurred during merging: {e}")

# Main
if selected_option == "Pricing Analysis Version - 2":
    pricing_analysis_v2()
else:
    st.title("Welcome to the Pricing Tool")
    st.write("Select 'Pricing Analysis Version - 2' from the sidebar to begin.")
