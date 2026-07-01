import os
import zipfile
from io import BytesIO

import pandas as pd
import streamlit as st


st.set_page_config(page_title="Excel / CSV Merger", layout="wide")

st.title("Excel / CSV Multi-Sheet Merger")


# -------------------------------------------------
# Large file support note
# -------------------------------------------------
st.info(
    "For 500 MB uploads, make sure `.streamlit/config.toml` has "
    "`maxUploadSize = 500` and `maxMessageSize = 500`."
)


# -------------------------------------------------
# Helpers
# -------------------------------------------------
def clean_sheet_name(name):
    invalid_chars = ["\\", "/", "*", "[", "]", ":", "?"]
    name = str(name).strip()

    for ch in invalid_chars:
        name = name.replace(ch, "_")

    return name[:31] if name else "Sheet1"


def normalize_columns(df):
    df.columns = [str(col).strip() for col in df.columns]
    return df


def read_csv_safely(uploaded_file):
    encodings = ["utf-8", "utf-8-sig", "latin1", "ISO-8859-1"]

    for encoding in encodings:
        try:
            uploaded_file.seek(0)
            return pd.read_csv(
                uploaded_file,
                encoding=encoding,
                dtype=str,
                low_memory=False
            )
        except Exception:
            continue

    uploaded_file.seek(0)
    return pd.read_csv(
        uploaded_file,
        dtype=str,
        encoding_errors="ignore",
        low_memory=False
    )


def get_file_sheets(uploaded_file):
    file_name = uploaded_file.name
    extension = os.path.splitext(file_name)[1].lower()

    if extension == ".csv":
        df = read_csv_safely(uploaded_file)
        df = normalize_columns(df)
        sheet_name = os.path.splitext(file_name)[0]
        return {sheet_name: df}

    if extension in [".xlsx", ".xls"]:
        uploaded_file.seek(0)
        excel = pd.ExcelFile(uploaded_file)

        sheets = {}

        for sheet in excel.sheet_names:
            uploaded_file.seek(0)
            df = pd.read_excel(
                uploaded_file,
                sheet_name=sheet,
                dtype=str
            )
            df = normalize_columns(df)
            sheets[sheet] = df

        return sheets

    return {}


def add_missing_columns(df1, df2):
    df1 = normalize_columns(df1.copy())
    df2 = normalize_columns(df2.copy())

    all_columns = list(dict.fromkeys(list(df1.columns) + list(df2.columns)))

    for col in all_columns:
        if col not in df1.columns:
            df1[col] = ""
        if col not in df2.columns:
            df2[col] = ""

    return df1[all_columns], df2[all_columns], all_columns


def append_merge(existing_df, new_df):
    existing_df, new_df, _ = add_missing_columns(existing_df, new_df)

    merged = pd.concat(
        [existing_df, new_df],
        ignore_index=True
    )

    return merged.fillna("")


def merge_using_common_column(existing_df, new_df, common_column):
    existing_df = normalize_columns(existing_df.copy())
    new_df = normalize_columns(new_df.copy())

    if common_column not in existing_df.columns:
        existing_df[common_column] = ""

    if common_column not in new_df.columns:
        new_df[common_column] = ""

    existing_df, new_df, all_columns = add_missing_columns(existing_df, new_df)

    existing_df[common_column] = existing_df[common_column].fillna("").astype(str).str.strip()
    new_df[common_column] = new_df[common_column].fillna("").astype(str).str.strip()

    merged = existing_df.merge(
        new_df,
        on=common_column,
        how="outer",
        suffixes=("", "_new")
    )

    for col in all_columns:
        if col == common_column:
            continue

        new_col = col + "_new"

        if new_col in merged.columns:
            merged[col] = merged[col].replace("", pd.NA)
            merged[col] = merged[col].fillna(merged[new_col])
            merged.drop(columns=[new_col], inplace=True)

    return merged.fillna("")


def create_excel_output(merged_sheets):
    output = BytesIO()

    with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
        used_sheet_names = set()

        for sheet_name, df in merged_sheets.items():
            safe_name = clean_sheet_name(sheet_name)
            original_name = safe_name
            counter = 1

            while safe_name in used_sheet_names:
                suffix = f"_{counter}"
                safe_name = original_name[:31 - len(suffix)] + suffix
                counter += 1

            used_sheet_names.add(safe_name)

            df.to_excel(
                writer,
                sheet_name=safe_name,
                index=False
            )

    output.seek(0)
    return output


def create_csv_output(df):
    return df.to_csv(index=False).encode("utf-8-sig")


def create_zip_csv_output(merged_sheets):
    zip_buffer = BytesIO()

    with zipfile.ZipFile(zip_buffer, "w", zipfile.ZIP_DEFLATED) as zip_file:
        for sheet_name, df in merged_sheets.items():
            csv_data = df.to_csv(index=False).encode("utf-8-sig")
            zip_file.writestr(f"{clean_sheet_name(sheet_name)}.csv", csv_data)

    zip_buffer.seek(0)
    return zip_buffer


# -------------------------------------------------
# Sidebar
# -------------------------------------------------
st.sidebar.header("Merge Settings")

merge_type = st.sidebar.radio(
    "Merge Type",
    [
        "Append rows",
        "Merge using common column"
    ]
)

common_column = ""

if merge_type == "Merge using common column":
    common_column = st.sidebar.text_input(
        "Common Column Name",
        value="Stock Number"
    ).strip()

merge_same_sheet_names = st.sidebar.checkbox(
    "Merge sheets with same name together",
    value=True
)

download_format = st.sidebar.radio(
    "Recommended download format",
    [
        "Excel workbook",
        "CSV zip"
    ],
    index=1
)

output_file_name = st.sidebar.text_input(
    "Output File Name",
    value="merged_output"
).strip()

if not output_file_name:
    output_file_name = "merged_output"


# -------------------------------------------------
# Upload
# -------------------------------------------------
uploaded_files = st.file_uploader(
    "Upload Excel or CSV files",
    type=["xlsx", "xls", "csv"],
    accept_multiple_files=True
)


# -------------------------------------------------
# Main
# -------------------------------------------------
if not uploaded_files:
    st.warning("Upload two or more Excel/CSV files to start.")
    st.stop()

if len(uploaded_files) < 2:
    st.warning("Please upload at least two files.")
    st.stop()

if merge_type == "Merge using common column" and not common_column:
    st.warning("Please enter the common column name.")
    st.stop()


merged_sheets = {}
report_rows = []

progress = st.progress(0)
status = st.empty()

total_files = len(uploaded_files)

for index, uploaded_file in enumerate(uploaded_files, start=1):
    status.write(f"Reading: **{uploaded_file.name}**")

    file_sheets = get_file_sheets(uploaded_file)

    for sheet_name, df in file_sheets.items():
        file_base_name = os.path.splitext(uploaded_file.name)[0]

        if merge_same_sheet_names:
            output_sheet_name = sheet_name
        else:
            output_sheet_name = f"{file_base_name}_{sheet_name}"

        rows_read = len(df)
        columns_read = len(df.columns)

        if output_sheet_name not in merged_sheets:
            if merge_type == "Merge using common column":
                if common_column not in df.columns:
                    df[common_column] = ""

            merged_sheets[output_sheet_name] = df.copy()

            rows_before = 0
            rows_after = len(df)

        else:
            rows_before = len(merged_sheets[output_sheet_name])

            if merge_type == "Append rows":
                merged_sheets[output_sheet_name] = append_merge(
                    merged_sheets[output_sheet_name],
                    df
                )
            else:
                merged_sheets[output_sheet_name] = merge_using_common_column(
                    merged_sheets[output_sheet_name],
                    df,
                    common_column
                )

            rows_after = len(merged_sheets[output_sheet_name])

        report_rows.append({
            "File": uploaded_file.name,
            "Input Sheet": sheet_name,
            "Output Sheet": output_sheet_name,
            "Rows Read": rows_read,
            "Columns Read": columns_read,
            "Rows Before Merge": rows_before,
            "Rows After Merge": rows_after
        })

    progress.progress(index / total_files)

status.empty()
progress.empty()

st.success("Files merged successfully!")


# -------------------------------------------------
# Preview
# -------------------------------------------------
st.subheader("Preview")

selected_sheet = st.selectbox(
    "Select sheet",
    list(merged_sheets.keys())
)

selected_df = merged_sheets[selected_sheet]

st.write(
    f"Rows: **{len(selected_df)}** | Columns: **{len(selected_df.columns)}**"
)

st.dataframe(
    selected_df.head(200),
    use_container_width=True
)


# -------------------------------------------------
# Downloads
# -------------------------------------------------
st.subheader("Download")

if download_format == "Excel workbook":
    excel_output = create_excel_output(merged_sheets)

    st.download_button(
        label="Download Merged Excel",
        data=excel_output,
        file_name=f"{output_file_name}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

else:
    zip_output = create_zip_csv_output(merged_sheets)

    st.download_button(
        label="Download Merged CSV ZIP",
        data=zip_output,
        file_name=f"{output_file_name}.zip",
        mime="application/zip"
    )


csv_output = create_csv_output(selected_df)

st.download_button(
    label="Download Selected Sheet as CSV",
    data=csv_output,
    file_name=f"{clean_sheet_name(selected_sheet)}.csv",
    mime="text/csv"
)


# -------------------------------------------------
# Merge Report
# -------------------------------------------------
st.subheader("Merge Report")

report_df = pd.DataFrame(report_rows)

st.dataframe(
    report_df,
    use_container_width=True
)

st.download_button(
    label="Download Merge Report CSV",
    data=create_csv_output(report_df),
    file_name="merge_report.csv",
    mime="text/csv"
)
