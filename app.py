import streamlit as st
import pandas as pd
import io
from openpyxl.styles import PatternFill, colors, Font, Border, Side, Alignment
from openpyxl.formatting.rule import ColorScaleRule
from openpyxl.utils import get_column_letter
import matplotlib.colors as mcolors

# App configuration
st.set_page_config(
    page_title="Option Chain Formatter",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded"
)

# Custom CSS for better styling
st.markdown("""
    <style>
    .main {
        max-width: 1000px;
        padding: 2rem;
    }
    .stButton>button {
        width: 100%;
        padding: 0.5rem;
        border-radius: 5px;
    }
    .stDownloadButton>button {
        background-color: #4CAF50;
        color: white;
    }
    </style>
""", unsafe_allow_html=True)

def process_option_chain(df):
    """
    Reads the input Excel (with the 2-row header pattern),
    selects & renames columns, computes CE/PE MONEY & BEP.
    """
    try:
        # 1) Read raw with no header so we can control header rows
        raw = df.copy()
        
        # Skip the first row (index 0) which contains "Exported data"
        # The actual headers are in the second row (index 1)
        headers = raw.iloc[1].tolist()  # Second row contains the real headers
        
        # Clean up headers (remove any extra whitespace and handle NaN values)
        headers = [str(h).strip() if pd.notna(h) and str(h).strip() != '' else f"Column_{i}" 
                  for i, h in enumerate(headers)]
        
        # Data starts from row 3 (index 2) because:
        # - Row 0: "Exported data" and empty cells
        # - Row 1: Headers
        # - Row 2: Data starts here
        data = raw.iloc[2:].copy()
        data.columns = headers
        data = data.reset_index(drop=True)  # Reset index after skipping rows

        # Helpers to handle duplicate column names (e.g., VWAP, OI, OI Chg appear twice)
        def col_positions(name: str):
            return [i for i, v in enumerate(headers) if str(v).strip() == name]

        def series_by_pos(name: str, occurrence: int = 0) -> pd.Series:
            pos_list = col_positions(name)
            if len(pos_list) <= occurrence:
                raise KeyError(f"Column '{name}' occurrence {occurrence} not found.\n"
                              f"Available positions for '{name}': {pos_list}")
            return data.iloc[:, pos_list[occurrence]]

        # 3) Map all needed columns (CE = first occurrence, PE = second)
        time_s = series_by_pos("Time", 0)
        strike_s = pd.to_numeric(series_by_pos("Strike Price", 0), errors="coerce")

        ce_oi_chg_s = pd.to_numeric(series_by_pos("OI Chg", 0), errors="coerce")
        ce_oi_s = pd.to_numeric(series_by_pos("OI", 0), errors="coerce")
        ce_vwap_s = pd.to_numeric(series_by_pos("VWAP", 0), errors="coerce")

        pe_vwap_s = pd.to_numeric(series_by_pos("VWAP", 1), errors="coerce")
        pe_oi_s = pd.to_numeric(series_by_pos("OI", 1), errors="coerce")
        pe_oi_chg_s = pd.to_numeric(series_by_pos("OI Chg", 1), errors="coerce")

        # 4) Build output with exact order + computed fields
        out = pd.DataFrame()
        out["Time"] = time_s
        out["CE OI CHANGE"] = ce_oi_chg_s
        out["CE OI"] = ce_oi_s
        out["CE MONEY"] = ((out["CE OI CHANGE"] * ce_vwap_s) / 10_000_000).round(0)
        out["CE BEP"] = (strike_s + ce_vwap_s).round(0)
        out["CE VWAP"] = ce_vwap_s
        out["Strike Price"] = strike_s
        out["PE VWAP"] = pe_vwap_s
        out["PE BEP"] = (strike_s - pe_vwap_s).round(0)
        out["PE MONEY"] = ((pe_oi_chg_s * pe_vwap_s) / 10_000_000).round(0)
        out["PE OI"] = pe_oi_s
        out["PE OI CHANGE"] = pe_oi_chg_s

        # Final column order
        final_order = [
            "Time",
            "CE OI CHANGE", "CE OI", "CE MONEY", "CE BEP", "CE VWAP",
            "Strike Price",
            "PE VWAP", "PE BEP", "PE MONEY", "PE OI", "PE OI CHANGE"
        ]
        
        return out[final_order]
        
    except Exception as e:
        import traceback
        st.error(f"Error processing file: {str(e)}")
        st.error(f"Error details: {traceback.format_exc()}")
        return None

def apply_formatting(worksheet, df):
    """Apply formatting to the Excel worksheet"""
    # Define border style
    thin_border = Border(
        left=Side(style='thin'),
        right=Side(style='thin'),
        top=Side(style='thin'),
        bottom=Side(style='thin')
    )
    
    # Format headers
    header_font = Font(bold=True, size=12, color='FFFFFF')
    header_fill = PatternFill(start_color='4F81BD', end_color='4F81BD', fill_type='solid')
    
    # Apply formatting to all cells
    for row in worksheet.iter_rows():
        for cell in row:
            cell.border = thin_border
            
            # Format header row
            if cell.row == 1:  # Header row
                cell.font = header_font
                cell.fill = header_fill
                cell.alignment = Alignment(horizontal='center', vertical='center')
            # Format data rows
            else:
                cell.alignment = Alignment(horizontal='right', vertical='center')
                
                # Right-align numeric columns
                if isinstance(cell.value, (int, float)):
                    cell.number_format = '#,##0'  # Format numbers with thousand separators
    
    # Color negative OI Change values red
    oi_change_cols = {'B': 'CE OI CHANGE', 'K': 'PE OI CHANGE'}
    for col_letter, col_name in oi_change_cols.items():
        if col_name in df.columns:
            col_idx = df.columns.get_loc(col_name) + 1  # +1 for 1-based index
            for row in range(2, len(df) + 2):  # +2 for header and 1-based index
                cell = worksheet.cell(row=row, column=col_idx)
                if cell.value is not None:
                    if cell.value < 0:
                        cell.font = Font(color='FF0000')  # Red for negative
                    
    # Add color scaling for OI Change columns (B and K)
    for col in ['B', 'L']:  # CE OI CHANGE and PE OI CHANGE
        color_scale_rule = ColorScaleRule(
            start_type='min', 
            start_color='9DC3E6',  # Light Blue for min
            mid_type='num', 
            mid_value=0, 
            mid_color='FFFFFF',     # White for zero
            end_type='max', 
            end_color='1F4E79'     # Dark Blue for max
        )
        worksheet.conditional_formatting.add(f'{col}2:{col}{len(df)+1}', color_scale_rule)
    
    # Add color scaling for OI columns (C and K)
    for col in ['C', 'K']:  # CE OI and PE OI
        color_scale_rule = ColorScaleRule(
            start_type='min', 
            start_color='F5F5DC',  # Light Beige for min
            mid_type='percentile', 
            mid_value=50, 
            mid_color='FFA07A',    # Orange for median
            end_type='max', 
            end_color='786C3B'     # Dark Brown for max
        )
        worksheet.conditional_formatting.add(f'{col}2:{col}{len(df)+1}', color_scale_rule)
    
    # Add color scaling for Money columns (D and J)
    for col in ['D', 'J']:  # CE MONEY and PE MONEY
        color_scale_rule = ColorScaleRule(
            start_type='min', 
            start_color='F8696B',  # Light Red for min
            mid_type='num', 
            mid_value=0, 
            mid_color='FFEB84',    # Yellow for zero
            end_type='max', 
            end_color='63BE7B'     # Green for max
        )
        worksheet.conditional_formatting.add(f'{col}2:{col}{len(df)+1}', color_scale_rule)
    
    # Auto-adjust column widths
    for column in worksheet.columns:
        max_length = 0
        column_letter = get_column_letter(column[0].column)
        for cell in column:
            try:
                if len(str(cell.value)) > max_length:
                    max_length = len(str(cell.value))
            except:
                pass
        adjusted_width = (max_length + 2) * 1.2
        worksheet.column_dimensions[column_letter].width = min(adjusted_width, 25)
    
    return worksheet

def style_dataframe_for_preview(df):
    """Apply color scaling to the dataframe for UI preview."""
    styled_df = df.style

    # Define color maps
    cmap_oi_change = mcolors.LinearSegmentedColormap.from_list("oi_change", ['#9DC3E6', '#FFFFFF', '#1F4E79'])
    cmap_oi = mcolors.LinearSegmentedColormap.from_list("oi", ['#FFF2CC', '#F4B183', '#8B4513'])
    cmap_money = mcolors.LinearSegmentedColormap.from_list("money", ['#F8696B', '#FFEB84', '#63BE7B'])

    # Apply styles
    columns_to_style = {
        'CE OI CHANGE': cmap_oi_change,
        'PE OI CHANGE': cmap_oi_change,
        'CE OI': cmap_oi,
        'PE OI': cmap_oi,
        'CE MONEY': cmap_money,
        'PE MONEY': cmap_money
    }

    for col, cmap in columns_to_style.items():
        if col in df.columns:
            styled_df = styled_df.background_gradient(cmap=cmap, subset=[col])
    
    return styled_df

def main():
    # App header
    st.title("📊 Option Chain Formatter")
    st.markdown("---")
    st.write("Upload your option chain Excel file to process and download the formatted version.")
    
    # File uploader
    uploaded_file = st.file_uploader("Choose an Excel file", type=["xlsx", "xls"])
    
    if uploaded_file is not None:
        try:
            # Read the uploaded file with header=None to handle the multi-row header
            # and ensure all data is read as strings to prevent type issues
            raw_df = pd.read_excel(uploaded_file, header=None, dtype=str)
            
            with st.spinner('Processing your file...'):
                # Process the data
                processed_df = process_option_chain(raw_df)
                
                if processed_df is not None:
                    # Create Excel file in memory
                    output = io.BytesIO()
                    with pd.ExcelWriter(output, engine='openpyxl') as writer:
                        processed_df.to_excel(writer, index=False, sheet_name='OptionChain')
                        worksheet = writer.sheets['OptionChain']
                        
                        # Apply formatting
                        apply_formatting(worksheet, processed_df)
                    
                    st.success("✅ Processing complete!")
                    
                    # Generate output filename
                    import os
                    input_filename = uploaded_file.name
                    base_name = os.path.splitext(input_filename)[0]
                    output_filename = f"{base_name}_processed.xlsx"
                    
                    # Download button
                    st.download_button(
                        label="⬇️ Download Processed File",
                        data=output.getvalue(),
                        file_name=output_filename,
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    )
                    
                    # Show preview of all rows
                    st.subheader("Preview of Processed Data")
                    # Set pandas display options to show all rows
                    pd.set_option('display.max_rows', None)
                    st.dataframe(style_dataframe_for_preview(processed_df))
                    # Reset the option after displaying
                    pd.reset_option('display.max_rows')
                    
        except Exception as e:
            st.error(f"❌ An error occurred: {str(e)}")
    
    # Sidebar with instructions
    with st.sidebar:
        st.title("ℹ️ Instructions")
        st.markdown("""
        1. Upload your option chain Excel file( iCharts file format ONLY is supported )
        2. Wait for processing to complete
        3. Download the formatted file or Preview on the main page
        """)
        
        st.markdown("---")
        st.markdown("### About")
        st.markdown("This tool helps you format option chain data to get BEPs and OI Money data.")
        st.markdown("© 2025 Manikanth Mulukutla. All rights reserved.")

if __name__ == "__main__":
    main()
