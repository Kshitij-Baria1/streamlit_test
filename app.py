import difflib
import shutil
from pathlib import Path

import pandas as pd
import streamlit as st
from st_aggrid import AgGrid, GridOptionsBuilder, GridUpdateMode
from streamlit import sidebar

import constants

st.set_page_config(layout="wide")

TEMP_EXCEL_ROOT_FOLDER = Path(__file__).parent / "temp_excel_files"

if "skipped_excel_files" not in st.session_state:
    st.session_state.skipped_excel_files = []


# normalize column name
def normalize_column_name(name):
    return str(name).strip().lower().replace(" ", "_")


# align column name
def align_column_names(actual_column_names, expected_column_names):
    aligned_columns = {}
    for expected_column_name in expected_column_names:
        match = difflib.get_close_matches(expected_column_name, actual_column_names, n=1, cutoff=0.7)
        if match:
            aligned_columns[match[0]] = expected_column_name
    return aligned_columns


# detect header row
@st.cache_data
def detect_header_row(file, expected_headers, max_rows=10):
    preview = pd.read_excel(file, engine="openpyxl", header=None, nrows=max_rows)
    expected_headers = [header.lower() for header in expected_headers]

    for i, row in preview.iterrows():
        row_clean = [str(cell).strip().lower() for cell in row if pd.notna(cell)]
        match_count = sum(
            1 for cell in row_clean
            for expected in expected_headers
            if difflib.get_close_matches(cell, [expected], cutoff=0.8)
        )
        if match_count >= len(expected_headers) * 0.3:
            return i
    return 0


# load and standardize excel
@st.cache_data
def load_and_standardize_excel(file_path: Path, expected_headers):
    header_row = detect_header_row(file_path, expected_headers)
    df = pd.read_excel(file_path, engine="openpyxl", header=header_row)

    # Normalize columns
    actual_normalized = [normalize_column_name(col) for col in df.columns]
    expected_normalized = [normalize_column_name(h) for h in expected_headers]

    aligned = align_column_names(actual_normalized, expected_normalized)

    # Rename matched columns
    df.columns = actual_normalized
    df = df.rename(columns=aligned)

    # Rename to original expected header format
    reverse_lookup = {normalize_column_name(h): h for h in expected_headers}
    df = df.rename(columns={col: reverse_lookup.get(col, col) for col in df.columns})

    # Add missing expected columns with NaN
    for expected in expected_headers:
        if expected not in df.columns:
            df[expected] = pd.NA

    # Reorder columns
    df = df[expected_headers]

    return df


# copy updated excel files
def copy_updated_excel_files(source_root_folder: Path, target_root_folder: Path):
    copied_files = []
    source_excel_files = [file for file in Path(source_root_folder).rglob("*.xlsx") if not file.name.startswith("~$")]

    for file in source_excel_files:
        rel_path = file.relative_to(source_root_folder)
        dest_path = target_root_folder / rel_path
        dest_path.parent.mkdir(parents=True, exist_ok=True)

        try:
            if not dest_path.exists() or file.stat().st_mtime > dest_path.stat().st_mtime:
                shutil.copy2(file, dest_path)
            copied_files.append(dest_path)
        except Exception as e:
            st.session_state.skipped_excel_files.append({
                "file": str(file),
                "reason": "copy_failed",
                "error": str(e)
            })
            st.warning(f"Failed to copy {file}: {e}")

    return copied_files


# clean stale temp files
def clean_stale_temp_files():
    for temp_file in TEMP_EXCEL_ROOT_FOLDER.rglob("*.xlsx"):
        original_file = constants.ORIGINAL_EXCEL_ROOT_FOLDER / temp_file.relative_to(TEMP_EXCEL_ROOT_FOLDER)
        if not original_file.exists():
            temp_file.unlink()


# ==== Main App ====

# add excel_files to session state
if 'excel_files' not in st.session_state:
    st.session_state.excel_files = copy_updated_excel_files(Path(constants.ORIGINAL_EXCEL_ROOT_FOLDER),
                                                            TEMP_EXCEL_ROOT_FOLDER)

# refresh data on button press
if st.button("🔄 Refresh Excel Data"):
    st.session_state.skipped_excel_files = []
    st.cache_data.clear()
    st.session_state.excel_files = copy_updated_excel_files(Path(constants.ORIGINAL_EXCEL_ROOT_FOLDER),
                                                            TEMP_EXCEL_ROOT_FOLDER)

# Select a file from the updated/existing files stored in session state
selected_file = st.selectbox("Select a file", st.session_state.excel_files)

# if file is selected
if selected_file:
    try:
        st.session_state.df = load_and_standardize_excel(selected_file, constants.EXPECTED_HEADERS)
        st.session_state.filtered_df = st.session_state.df.copy()
    except Exception as e:
        st.session_state.skipped_excel_files.append({
            "file": str(selected_file),
            "reason": "exception",
            "error": str(e)
        })
        st.warning(f"Could not load {selected_file.name}: {e}")

# add filter fields in session state
if "license_name_input" not in st.session_state:
    st.session_state.license_name_input = ""
if "entity_jurisdiction_multiselect" not in st.session_state:
    st.session_state.entity_jurisdiction_multiselect = []
if "license_jurisdiction_name" not in st.session_state:
    st.session_state.license_jurisdiction_name = ""


# reset filters
def reset_filters():
    st.session_state.license_name_input = ""
    st.session_state.entity_jurisdiction_multiselect = []
    st.session_state.license_jurisdiction_name = ""

    st.session_state.filtered_df = st.session_state.df.copy()


# filter sidebar
with sidebar:
    st.title("Filter")
    # filter fields
    license_name = st.text_input("License Name", key="license_name_input")
    entity_jurisdiction = st.multiselect("Entity jurisdiction", options=constants.USA_STATE_ABBREVIATIONS,
                                         key="entity_jurisdiction_multiselect")
    license_jurisdiction_name = st.text_input("License jurisdiction name", key="license_jurisdiction_name")

    # reset filters button
    reset_filters_button = st.button("❌ Reset Filters", on_click=reset_filters)

# apply filters
if license_name:
    st.session_state.filtered_df = st.session_state.filtered_df[
        st.session_state.filtered_df["License Name"].str.contains(license_name, case=False, na=False)]
if entity_jurisdiction:
    st.session_state.filtered_df = st.session_state.filtered_df[
        st.session_state.filtered_df["Entity Jurisdiction"].isin(entity_jurisdiction)]
if license_jurisdiction_name:
    st.session_state.filtered_df = st.session_state.filtered_df[
        st.session_state.filtered_df["License jurisdiction name"].str.contains(license_jurisdiction_name, case=False,
                                                                               na=False)]

# add selected_rows_df to session state
if "selected_rows_df" not in st.session_state:
    st.session_state.selected_rows_df = pd.DataFrame()

# display filtered_df
if "filtered_df" in st.session_state and not st.session_state.filtered_df.empty:
    # add row button
    if st.button("➕ Add row"):
        empty_row = pd.DataFrame([{col: "" for col in st.session_state.filtered_df.columns}])
        st.session_state.filtered_df = pd.concat([st.session_state.filtered_df, empty_row], ignore_index=True)

    # show filtered_df
    gb = GridOptionsBuilder.from_dataframe(st.session_state.filtered_df)
    gb.configure_default_column(editable=True)
    gb.configure_selection(selection_mode="multiple", header_checkbox=True, use_checkbox=True)
    gb.configure_pagination(paginationAutoPageSize=False, paginationPageSize=50)
    grid_response = AgGrid(
        st.session_state.filtered_df,
        gridOptions=gb.build(),
        update_mode=GridUpdateMode.VALUE_CHANGED,
        use_container_width=True,
        height=300
    )

    # get selected rows from grid response
    selected_rows = pd.DataFrame(grid_response["selected_rows"])

    # columns to add and clear selection
    col1, col2, col3 = st.columns(3)
    with col1:
        if st.button("💾 Save Changes"):
            try:
                # grab the truly edited data
                edited_df = pd.DataFrame(grid_response["data"])
                # clean up any stray metadata
                edited_df = edited_df[[c for c in edited_df.columns if c in constants.EXPECTED_HEADERS]]
                # re-order to your expected schema
                edited_df = edited_df[constants.EXPECTED_HEADERS]

                # resolve the original path
                original_path = (
                        constants.ORIGINAL_EXCEL_ROOT_FOLDER
                        / selected_file.relative_to(TEMP_EXCEL_ROOT_FOLDER)
                )

                # write it out
                try:
                    edited_df.to_excel(original_path, index=False, engine="openpyxl")
                    st.toast(f"Saved to original {original_path} with your edits.", icon="✅")
                except Exception as e:
                    st.toast(f"Could not save edits to original file: {e}", icon="⚠️")

                # (optional) Mirror to temp so the UI reloads
                try:
                    selected_file.write_bytes(original_path.read_bytes())
                    st.toast(f"Saved to temp {selected_file} with your edits.", icon="✅")
                except Exception as e:
                    st.toast(f"Could not save edits to temp file: {e}", icon="⚠️")

                # clear & reload only this file’s cache
                detect_header_row.clear()
                load_and_standardize_excel.clear()
                st.session_state.df = load_and_standardize_excel(selected_file, constants.EXPECTED_HEADERS)
                st.session_state.filtered_df = st.session_state.df.copy()
            except Exception as e:
                st.toast(f"Could not save edits: {e}", icon="⚠️")
    with col2:
        if st.button("➕ Add to selection") and not selected_rows.empty:
            selected_rows = selected_rows[st.session_state.filtered_df.columns]
            key_columns = [col for col in st.session_state.filtered_df.columns if col != "Entity Name"]
            st.session_state.selected_rows_df = pd.concat(
                [st.session_state.selected_rows_df, selected_rows],
                ignore_index=True
            ).drop_duplicates(key_columns)
    with col3:
        if st.button("🗑️ Clear selection"):
            st.session_state.selected_rows_df = pd.DataFrame()
            selected_rows = []

    if not st.session_state.selected_rows_df.empty:
        col1, col2 = st.columns(2)

        if "entity_name_input" not in st.session_state:
            st.session_state.entity_name_input = ""

        with col1:
            entity_name = st.text_input("Entity name", key="entity_name_input")
        with col2:
            if st.button("Add entity name"):
                if st.session_state.entity_name_input:
                    if entity_name.strip():
                        st.session_state.selected_rows_df["Entity Name"] = entity_name

        st.dataframe(st.session_state.selected_rows_df)

        # Text area for TSV
        tsv = st.session_state.selected_rows_df.to_csv(index=False, sep="\t", header=False)
        st.text_area("📋 Copy the table below (TSV format)", tsv, height=250, key="tsv_output")
