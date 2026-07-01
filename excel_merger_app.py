import os
from io import BytesIO

import pandas as pd
import streamlit as st


st.set_page_config(page_title="File Merger", layout="wide")
st.title("Excel / CSV File Merger")


uploaded_files = st.file_uploader(
    "Upload Excel or CSV files",
    type=["xlsx", "xls", "csv"],
    accept_multiple_files=True
)


COLUMN_ALIASES = {
    "stock num": "Stock Number",
    "stock_num": "Stock Number",
    "stock number": "Stock Number",
    "stock no": "Stock Number",
    "stock_no": "Stock Number",
    "sku": "Stock Number",

    "image url": "Image URL",
    "image_url": "Image URL",
    "image": "Image URL",

    "price": "Price",
    "total price": "Price",
    "total_price": "Price",

    "description": "Description",
    "desc": "Description",

    "title": "Title",
    "name": "Title",
    "product name": "Title",
}


def normalize_column_name(col):
    col = str(col).strip()
    key = col.lower().strip().replace("-", " ").replace("_", " ")
    key = " ".join(key.split())

    return COLUMN_ALIASES.get(key, col)


def normalize_columns(df):
    df = df.copy()

    df.columns = [normalize_column_name(col) for col in df.columns]

    # Merge duplicate columns into one
    final_df = pd.DataFrame()

    for col in df.columns.unique():
        same_cols = df.loc[:, df.columns == col]

        if same_cols.shape[1] == 1:
            final_df[col] = same_cols.iloc[:, 0]
        else:
            final_df[col] = same_cols.bfill(axis=1).iloc[:, 0]

    return final_df.fillna("")


def read_file(file):
    ext = os.path.splitext(file.name)[1].lower()

    if ext == ".csv":
        df = pd.read_csv(file, dtype=str).fillna("")
        return [normalize_columns(df)]

    sheets = pd.read_excel(file, sheet_name=None, dtype=str)

    all_sheets = []
    for sheet_name, df in sheets.items():
        df = normalize_columns(df)
        all_sheets.append(df)

    return all_sheets


def merge_files(uploaded_files):
    all_data = []

    for file in uploaded_files:
        file_dfs = read_file(file)

        for df in file_dfs:
            all_data.append(df)

    merged_df = pd.concat(all_data, ignore_index=True, sort=False)
    merged_df = normalize_columns(merged_df)

    return merged_df.fillna("")


def create_excel(df):
    output = BytesIO()

    with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
        df.to_excel(writer, index=False, sheet_name="Merged Data")

    output.seek(0)
    return output


if uploaded_files:
    merged_df = merge_files(uploaded_files)

    st.success("Files merged successfully!")

    st.write(f"Total Rows: **{len(merged_df)}**")
    st.write(f"Total Columns: **{len(merged_df.columns)}**")

    st.dataframe(merged_df.head(200), use_container_width=True)

    excel_file = create_excel(merged_df)

    st.download_button(
        label="Download Merged Excel",
        data=excel_file,
        file_name="merged_output.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

else:
    st.info("Upload Excel or CSV files to merge.")
