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

def extract_data_by_reference(file):
    in_memory_file = BytesIO(file.read())
    wb = load_workbook(in_memory_file, data_only=False)
    ws = wb.active
    rows = list(ws.iter_rows(values_only=False))

    # Capture title row with merged cells
    merged_ranges = ws.merged_cells.ranges
    title_row = []
    for cell in rows[0]:
        title_row.append({
            "value": cell.value,
            "fill": copy(cell.fill),
            "font": copy(cell.font),
            "alignment": copy(cell.alignment),
            "border": copy(cell.border)
        })

    # Extract column headers (row 1)
    headers = [cell.value for cell in rows[1]]
    ref_col_idx = next((idx for idx, val in enumerate(headers) if val and "reference" in str(val).lower()), None)
    if ref_col_idx is None:
        raise ValueError("Reference Number column not found")

    header_cells = []
    for cell in rows[1]:
        header_cells.append({
            "value": cell.value,
            "fill": copy(cell.fill),
            "font": copy(cell.font),
            "alignment": copy(cell.alignment),
            "border": copy(cell.border)
        })

    # Extract formatted data
    formatted_data = {}
    for row in rows[2:]:
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

    return title_row, header_cells, ref_col_idx, formatted_data, merged_ranges

def merge_files_by_reference(uploaded_files):
    all_references = set()
    file_data = []
    title_rows = []

    for file in uploaded_files:
        title_row, headers, ref_idx, data, merged_ranges = extract_data_by_reference(file)
        title_rows.append((title_row, merged_ranges))
        file_data.append((headers, ref_idx, data))
        all_references.update(data.keys())

    return sorted(all_references), file_data, title_rows

def apply_border(cell):
    thin = Side(style="thin")
    cell.border = Border(top=thin, bottom=thin, left=thin, right=thin)

def create_workbook(all_refs, file_data, title_rows):
    wb = Workbook()
    ws = wb.active
    ws.title = "Merged"

    current_row = 1

    # Write reference number header
    ws.cell(row=current_row + 1, column=1, value="Reference Number")
    ref_header_style = file_data[0][0][file_data[0][1]]
    cell = ws.cell(row=current_row + 1, column=1)
    cell.fill = ref_header_style["fill"]
    cell.font = ref_header_style["font"]
    cell.alignment = ref_header_style["alignment"]
    apply_border(cell)

    col_offset = 2  # start from column 2 since column 1 is for Reference Number

    # --- For each file, write its title row, headers, and track column positions ---
    column_positions = []

    for (title_row, merged_ranges), (headers, ref_idx, data) in zip(title_rows, file_data):
        start_col = col_offset
        num_cols = len([i for i in range(len(headers)) if i != ref_idx])

        # Title row (row 1) with merged cells
        for mrange in merged_ranges:
            if mrange.min_row == 1:
                start_merge = start_col
                end_merge = start_col + num_cols - 1
                merge_range = f"{ws.cell(row=current_row, column=start_merge).coordinate}:{ws.cell(row=current_row, column=end_merge).coordinate}"
                ws.merge_cells(merge_range)

        # Apply merged range for title
        merge_start = col_offset
        merge_end = col_offset + num_cols - 1
        merge_range = f"{ws.cell(row=current_row, column=merge_start).coordinate}:{ws.cell(row=current_row, column=merge_end).coordinate}"
        ws.merge_cells(merge_range)

        # Only write value and style to the top-left cell
        title_cell = title_row[0]  # First cell contains the main title text and style
        cell = ws.cell(row=current_row, column=merge_start, value=title_cell["value"])
        cell.fill = title_cell["fill"]
        cell.font = title_cell["font"]
        cell.alignment = title_cell["alignment"]
        apply_border(cell)


        # Header row (row 2)
        for i, cell_obj in enumerate(headers):
            if i != ref_idx:
                cell = ws.cell(row=current_row + 1, column=col_offset, value=cell_obj["value"])
                cell.fill = cell_obj["fill"]
                cell.font = cell_obj["font"]
                cell.alignment = cell_obj["alignment"]
                apply_border(cell)
                col_offset += 1

        column_positions.append((start_col, col_offset - 1, ref_idx, data))

    # Add final columns
    for i, col_name in enumerate(["Annual Volume/100", "x"]):
        cell = ws.cell(row=current_row + 1, column=col_offset + i, value=col_name)
        cell.font = Font(bold=True)
        cell.alignment = Alignment(horizontal="center", vertical="center")
        apply_border(cell)

    current_row += 2  # Move to the data rows

    # --- Write data rows ---
    for ref_val in all_refs:
        ws.cell(row=current_row, column=1, value=ref_val)
        apply_border(ws.cell(row=current_row, column=1))

        col_idx = 2
        for start_col, end_col, ref_idx, data in column_positions:
            if ref_val in data:
                row_data = data[ref_val]
                for i, cell_obj in enumerate(row_data):
                    if i == ref_idx:
                        continue
                    cell = ws.cell(row=current_row, column=col_idx, value=cell_obj["value"])
                    cell.fill = cell_obj["fill"]
                    cell.font = cell_obj["font"]
                    cell.alignment = cell_obj["alignment"]
                    apply_border(cell)
                    col_idx += 1
            else:
                for i in range(end_col - start_col + 1):
                    cell = ws.cell(row=current_row, column=col_idx, value=None)
                    apply_border(cell)
                    col_idx += 1

        # Empty final columns
        ws.cell(row=current_row, column=col_idx, value=None)
        apply_border(ws.cell(row=current_row, column=col_idx))
        ws.cell(row=current_row, column=col_idx + 1, value=None)
        apply_border(ws.cell(row=current_row, column=col_idx + 1))

        current_row += 1

    return wb


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

def pricing_analysis_v2():
    st.title("Pricing Analysis Version - 2 (All Headings + Merged Format)")

    uploaded_files = upload_files()
    if uploaded_files and st.button("Merge Files"):
        try:
            all_refs, file_data, title_rows = merge_files_by_reference(uploaded_files)
            wb = create_workbook(all_refs, file_data, title_rows)
            st.success("Files merged successfully with all headings and formatting.")
            get_download_link_with_styles(wb)
        except Exception as e:
            st.error(f"An error occurred during merging: {e}")

if selected_option == "Pricing Analysis Version - 2":
    pricing_analysis_v2()
else:
    st.title("Welcome to the Pricing Tool")
    st.write("Select 'Pricing Analysis Version - 2' from the sidebar to begin.")
