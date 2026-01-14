import streamlit as st
import pandas as pd
import json
import os
from io import BytesIO
import sys

st.set_page_config(page_title="Smart Excel Merger", layout="wide")
st.title("📊 Smart Excel Merger with Column Mapping")

BASE_DIR = os.path.dirname(sys.executable if getattr(sys, 'frozen', False) else __file__)
MAPPING_DIR = os.path.join(BASE_DIR, "mappings")

DATA_DIR = os.path.join(BASE_DIR, "data")
os.makedirs(DATA_DIR, exist_ok=True)

TEMPLATE_PATH = os.path.join(DATA_DIR, "template.xlsx")
os.makedirs(MAPPING_DIR, exist_ok=True)

st.subheader("📌 Template File")

template_exists = os.path.exists(TEMPLATE_PATH)

if template_exists:
    st.success("✅ Template file is set")
    if st.button("🔄 Replace Template"):
        os.remove(TEMPLATE_PATH)
        st.experimental_rerun()
else:
    template_file = st.file_uploader(
        "Upload Template Excel (one-time)",
        type=["xlsx"]
    )

    if template_file:
        with open(TEMPLATE_PATH, "wb") as f:
            f.write(template_file.getbuffer())
        st.success("Template saved successfully!")
        st.experimental_rerun()

if not os.path.exists(TEMPLATE_PATH):
    st.info("Please upload the template file to continue.")
    st.stop()

# ------------------------------------
# File Upload
# ------------------------------------
file2 = st.file_uploader("Upload Excel File B", type=["xlsx"], key="file2")

if not (template_file and file2):
    st.info("Upload both Excel files to continue.")
    st.stop()

# Wrap everything in a try block or use if/else properly
try:
        excel1 = pd.ExcelFile(template_file)
        excel2 = pd.ExcelFile(file2)
        
        # ------------------------------------
        # Sheet Selection
        # ------------------------------------
        col1, col2 = st.columns(2)

        with col1:
                sheet1 = st.selectbox("Select sheet from File A", excel1.sheet_names)

        with col2:
                sheet2 = st.selectbox("Select sheet from File B", excel2.sheet_names)

        df1 = pd.read_excel(excel1, sheet_name=sheet1)
        df2 = pd.read_excel(excel2, sheet_name=sheet2)

        # ------------------------------------
        # Row Count (Before)
        # ------------------------------------
        st.subheader("📈 Row Count (Before Merge)")
        c1, c2 = st.columns(2)
        c1.metric("File A Rows", len(df1))
        c2.metric("File B Rows", len(df2))

        # ------------------------------------
        # Load Saved Mapping
        # ------------------------------------
        st.subheader("💾 Load Column Mapping (Optional)")

        mapping_files = [f for f in os.listdir(MAPPING_DIR) if f.endswith(".json")]
        selected_mapping = st.selectbox(
        "Select saved mapping",
        ["None"] + mapping_files
        )

        column_mapping = {}

        if selected_mapping != "None":
                try:
                        with open(os.path.join(MAPPING_DIR, selected_mapping)) as f:
                                column_mapping = json.load(f)
                        st.success(f"Loaded mapping: {selected_mapping}")
                except Exception as e:
                        st.error(f"Error loading mapping: {e}")
        # ------------------------------------
        # Column Mapping UI
        # ------------------------------------
        st.subheader("🧩 Column Mapping (File A → File B)")

        # Add merge key selection
        st.info("Select which columns to use as merge keys (must map to same columns in both files)")

        mapped_columns = {}
        merge_keys = []

        for col in df1.columns:
                col_container = st.container()
                with col_container:
                        col_map, col_key = st.columns([3, 1])
                        
                with col_map:
                        default_value = column_mapping.get(col, "Ignore")
                        selected = st.selectbox(
                                f"{col} maps to",
                                ["Ignore"] + list(df2.columns),
                                index=(["Ignore"] + list(df2.columns)).index(default_value)
                                if default_value in ["Ignore"] + list(df2.columns) else 0,
                                key=f"map_{col}"
                        )
                
                with col_key:
                        if selected != "Ignore":
                                is_key = st.checkbox("Merge Key", key=f"key_{col}")
                                mapped_columns[col] = selected
                                if is_key:
                                        merge_keys.append(selected)

        if not mapped_columns:
                st.warning("Please map at least one column.")
                st.stop()

        if not merge_keys:
                st.warning("Please select at least one merge key column.")
                st.stop()

        # ------------------------------------
        # Save Mapping
        # ------------------------------------
        st.subheader("💾 Save Mapping Configuration")

        mapping_name = st.text_input("Mapping name (e.g. customer_merge)")

        if st.button("Save Mapping"):
                if mapping_name:
                        try:
                                mapping_data = {
                                        "column_mapping": mapped_columns,
                                        "merge_keys": merge_keys
                                }
                                with open(f"{MAPPING_DIR}/{mapping_name}.json", "w") as f:
                                        json.dump(mapping_data, f, indent=2)
                                st.success("Mapping saved successfully!")
                        except Exception as e:
                                st.error(f"Error saving mapping: {e}")
        else:
                st.error("Please enter a mapping name.")

        # ------------------------------------
        # Merge Options
        # ------------------------------------
        st.subheader("🔀 Merge Options")

        join_type = st.selectbox(
        "Merge type",
        ["inner", "left", "right", "outer"],
        help="inner: only matching rows | left: all from A | right: all from B | outer: all rows"
        )

        # ------------------------------------
        # Perform Merge
        # ------------------------------------
        if st.button("🚀 Merge Files"):
                try:
                        # Rename columns in df1 to match df2
                        df1_renamed = df1.rename(columns=mapped_columns)
                        
                        # Perform merge
                        merged_df = pd.merge(
                        df1_renamed,
                        df2,
                        on=merge_keys,
                        how=join_type,
                        suffixes=('_A', '_B')
                        )

                        # ------------------------------------
                        # Row Count (After)
                        # ------------------------------------
                        st.subheader("📊 Row Count (After Merge)")
                        m1, m2, m3 = st.columns(3)
                        m1.metric("File A Rows", len(df1))
                        m2.metric("File B Rows", len(df2))
                        m3.metric("Merged Rows", len(merged_df))

                        # ------------------------------------
                        # Preview
                        # ------------------------------------
                        st.subheader("✅ Merged Preview")
                        st.dataframe(merged_df.head(50))

                        # ------------------------------------
                        # Download
                        # ------------------------------------
                        # Use BytesIO instead of writing to disk
                        output = BytesIO()
                        merged_df.to_excel(output, index=False, engine='openpyxl')
                        output.seek(0)

                        st.download_button(
                        "⬇️ Download Merged Excel",
                        output,
                        file_name="merged_output.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                        )
                
                except Exception as e:
                        st.error(f"❌ Merge failed: {e}")
                        st.exception(e)
    
except Exception as e:
    st.error(f"❌ Failed to read Excel file: {e}")
    st.stop()

