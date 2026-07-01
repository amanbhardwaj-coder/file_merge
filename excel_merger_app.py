import pandas as pd
import streamlit as st
from io import BytesIO

st.set_page_config(page_title="Excel Multi-Sheet Merger", layout="wide")

st.title("Excel Multi-Sheet Merger")

uploaded_files = st.file_uploader(
    "Upload two or more Excel files",
    type=["xlsx", "xls"],
    accept_multiple_files=True
)

merge_type = st.radio(
    "Merge Type",
    [
        "Append rows sheet-wise",
        "Merge using common column"
    ]
)

common_column = None
if merge_type == "Merge using common column":
    common_column = st.text_input(
        "Common column name",
        value="Stock Number"
    )

output_file_name = st.text_input(
    "Output file name",
    value="merged_output.xlsx"
)

def clean_sheet_name(name):
    invalid_chars = ["\\", "/", "*", "[", "]", ":", "?"]
    for ch in invalid_chars:
        name = name.replace(ch, "_")
    return name[:31]

def add_missing_columns(df1, df2):
    all_columns = list(dict.fromkeys(list(df1.columns) + list(df2.columns)))

    for col in all_columns:
        if col not in df1.columns:
            df1[col] = ""
        if col not in df2.columns:
            df2[col] = ""

    return df1[all_columns], df2[all_columns], all_columns

def append_merge(existing, new_df):
    existing, new_df, all_columns = add_missing_columns(existing, new_df)
    return pd.concat([existing, new_df], ignore_index=True)

def common_column_merge(existing, new_df, common_column):
    if common_column not in existing.columns:
        existing[common_column] = ""

    if common_column not in new_df.columns:
        new_df[common_column] = ""

    existing, new_df, all_columns = add_missing_columns(existing, new_df)

    merged = existing.merge(
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

    return merged

def create_output_excel(merged_sheets):
    output = BytesIO()

    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        for sheet_name, df in merged_sheets.items():
            df.to_excel(
                writer,
                sheet_name=clean_sheet_name(sheet_name),
                index=False
            )

    output.seek(0)
    return output

if uploaded_files:
    if len(uploaded_files) < 2:
        st.warning("Please upload at least two Excel files.")
    else:
        merged_sheets = {}

        for file in uploaded_files:
            st.write(f"Reading: **{file.name}**")

            excel = pd.ExcelFile(file)

            for sheet in excel.sheet_names:
                df = pd.read_excel(file, sheet_name=sheet)

                if sheet not in merged_sheets:
                    merged_sheets[sheet] = df.copy()
                else:
                    if merge_type == "Append rows sheet-wise":
                        merged_sheets[sheet] = append_merge(
                            merged_sheets[sheet],
                            df
                        )
                    else:
                        merged_sheets[sheet] = common_column_merge(
                            merged_sheets[sheet],
                            df,
                            common_column
                        )

        st.success("Files merged successfully!")

        st.subheader("Preview")

        selected_sheet = st.selectbox(
            "Select sheet to preview",
            list(merged_sheets.keys())
        )

        st.dataframe(merged_sheets[selected_sheet].head(100))

        output_excel = create_output_excel(merged_sheets)

        st.download_button(
            label="Download Merged Excel",
            data=output_excel,
            file_name=output_file_name,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
