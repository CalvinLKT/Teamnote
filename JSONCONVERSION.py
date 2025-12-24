import streamlit as st
import pandas as pd
import json
import re
from io import BytesIO

st.set_page_config(page_title="JSON Converter", layout="centered")
st.title("JSON Converter")
st.markdown("Upload multiple JSON files to convert them into a clean, readable Excel file.")

# Helper functions to extract numeric values
def extract_weight_kg(weight_str):
    if not weight_str:
        return None
    # Remove commas and find the number
    match = re.search(r"([\d,.]+)", str(weight_str).replace(",", ""))
    return float(match.group(1)) if match else None

def extract_volume_cbm(volume_str):
    if not volume_str:
        return None
    match = re.search(r"([\d,.]+)", str(volume_str).replace(",", ""))
    return float(match.group(1)) if match else None

# Helper function to extract clean filename (between last _ and .json)
def extract_clean_filename(original_filename):
    if original_filename.endswith(".json"):
        base = original_filename[:-5]  # remove .json
        if "_" in base:
            return base.split("_")[-1]
    return original_filename  # fallback

# File uploader
uploaded_files = st.file_uploader(
    "Choose JSON files",
    type=["json"],
    accept_multiple_files=True,
    help="Upload one or more JSON files with the shipping data structure."
)

if uploaded_files:
    all_rows = []

    for uploaded_file in uploaded_files:
        original_filename = uploaded_file.name
        clean_filename = extract_clean_filename(original_filename)  # e.g., EGLV235501600325

        try:
            data = json.load(uploaded_file)

            # Extract header fields
            swb_no = data.get("SWB-No.", "")
            port_loading = data.get("Port of Loading", "")
            port_discharge = data.get("Port of Discharge", "")
            port_delivery = data.get("Port of Delivery", "")

            # Extract Total fields
            total_cartons = data.get("Total", {}).get("Cartons", "")
            total_weight_str = data.get("Total", {}).get("Weight", "")
            total_volume_str = data.get("Total", {}).get("Volume", "")
            total_weight_kg = extract_weight_kg(total_weight_str)
            total_volume_cbm = extract_volume_cbm(total_volume_str)

            # Process only "Items" (skip "Subtotal" if exists)
            items = data.get("Items", [])

            for item in items:
                row = {
                    "Filename": clean_filename,  # Updated: clean name only
                    "SWB-No.": swb_no,
                    "Port of Loading": port_loading,
                    "Port of Discharge": port_discharge,
                    "Port of Delivery": port_delivery,
                    "Container No.": item.get("Container No.", ""),
                    "Seal No.": item.get("Seal No.", ""),
                    "Container Size": item.get("Container Size", ""),
                    "Cartons": item.get("Cartons", ""),
                    "Weight": item.get("Weight", ""),
                    "Weight (KGS)": extract_weight_kg(item.get("Weight", "")),
                    "Volume": item.get("Volume", ""),
                    "Volume (CBM)": extract_volume_cbm(item.get("Volume", "")),
                    "Total Cartons": total_cartons,
                    "Total Weight": total_weight_str,
                    "Total Weight (KGS)": total_weight_kg,
                    "Total Volume": total_volume_str,
                    "Total Volume (CBM)": total_volume_cbm,
                }
                all_rows.append(row)

        except json.JSONDecodeError:
            st.error(f"Invalid JSON in file: {original_filename}")
        except Exception as e:
            st.error(f"Error processing {original_filename}: {str(e)}")

    if all_rows:
        df = pd.DataFrame(all_rows)

        # Column order as previously agreed
        column_order = [
            "Filename",
            "SWB-No.",
            "Port of Loading",
            "Port of Discharge",
            "Port of Delivery",
            "Container No.",
            "Seal No.",
            "Container Size",
            "Cartons",
            "Weight",
            "Weight (KGS)",
            "Volume",
            "Volume (CBM)",
            "Total Cartons",
            "Total Weight",
            "Total Weight (KGS)",
            "Total Volume",
            "Total Volume (CBM)",
        ]
        df = df[column_order]

        # Create Excel in memory
        output = BytesIO()
        with pd.ExcelWriter(output, engine="openpyxl") as writer:
            df.to_excel(writer, index=False, sheet_name="Shipping Data")

        output.seek(0)

        st.success(f"Successfully processed {len(uploaded_files)} file(s) → {len(df)} rows generated.")

        st.download_button(
            label="Download Excel File",
            data=output,
            file_name="Shipping_Data_Converted.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )

        # Preview
        with st.expander("Preview the resulting table"):
            st.dataframe(df, use_container_width=True)
    else:
        st.warning("No valid data found in the uploaded files.")