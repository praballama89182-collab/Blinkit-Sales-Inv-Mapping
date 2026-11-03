import io
import pandas as pd
import streamlit as st
import openpyxl
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
from openpyxl.utils import get_column_letter

st.set_page_config(
    page_title="Quick-Commerce Inventory & Sales Analytics Dashboard",
    page_icon="📦",
    layout="wide",
    initial_sidebar_state="expanded"
)

# Custom Styling & Palette
st.markdown("""
    <style>
    .main { background-color: #F8F9FA; }
    .stMetric {
        background-color: #FFFFFF;
        padding: 15px;
        border-radius: 8px;
        box-shadow: 0 2px 4px rgba(0,0,0,0.05);
        border-top: 3px solid #1F4E78;
    }
    </style>
""", unsafe_allow_html=True)

st.title("📦 Quick-Commerce Inventory & Sales Analytics Dashboard")
st.markdown("Upload **Stock On Hand (Inventory)** and **Sales Summary** Excel reports to analyze sellable, warehouse, and darkstore stock levels, variation-level UoM metrics, and state-wise performance.")

# Sidebar Uploads
st.sidebar.header("📁 Data Upload Center")
inv_file = st.sidebar.file_uploader("Upload Inventory Report (.xlsx)", type=["xlsx", "xls"])
sales_file = st.sidebar.file_uploader("Upload Sales Summary Report (.xlsx)", type=["xlsx", "xls"])

@st.cache_data
def load_inventory_data(file):
    # Dynamically detect header row containing 'Item ID'
    df_raw = pd.read_excel(file, sheet_name=0, header=None)
    header_idx = 0
    for idx, row in df_raw.iterrows():
        row_vals = [str(val).strip().lower() for val in row.values if pd.notna(val)]
        if "item id" in row_vals:
            header_idx = idx
            break
            
    df = pd.read_excel(file, sheet_name=0, header=header_idx)
    df = df.loc[:, ~df.columns.str.contains('^Unnamed')]
    
    # Standardize column names
    rename_dict = {}
    for col in df.columns:
        c_lower = str(col).strip().lower()
        if "item id" in c_lower:
            rename_dict[col] = "Item ID"
        elif "item name" in c_lower:
            rename_dict[col] = "Item Name"
        elif "brand" in c_lower:
            rename_dict[col] = "Brand Name"
        elif "upc" in c_lower:
            rename_dict[col] = "UPC"
        elif "uom" in c_lower:
            rename_dict[col] = "UoM"
        elif "warehouse facility name" in c_lower or "warehouse name" in c_lower:
            rename_dict[col] = "Warehouse Facility Name"
        elif "total sellable" in c_lower:
            rename_dict[col] = "Total Sellable"
        elif c_lower == "warehouse":
            rename_dict[col] = "Warehouse"
        elif "darkstore" in c_lower:
            rename_dict[col] = "Darkstore"
            
    df = df.rename(columns=rename_dict)
    return df

@st.cache_data
def load_sales_data(file):
    df = pd.read_excel(file, sheet_name=0)
    if "Item Id" in df.columns:
        df = df.rename(columns={"Item Id": "Item ID"})
    return df

def style_excel_sheet(ws, header_bg_color="1F4E78"):
    header_fill = PatternFill(start_color=header_bg_color, end_color=header_bg_color, fill_type="solid")
    header_font = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
    cell_font = Font(name="Calibri", size=10)
    center_align = Alignment(horizontal="center", vertical="center", wrap_text=True)
    thin_side = Side(style='thin', color='D3D3D3')
    cell_border = Border(left=thin_side, right=thin_side, top=thin_side, bottom=thin_side)
    
    # Format Header Row
    for cell in ws[1]:
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = center_align
        cell.border = cell_border
        
    # Format Data Rows
    for row in ws.iter_rows(min_row=2, max_row=ws.max_row, min_col=1, max_col=ws.max_column):
        for cell in row:
            cell.font = cell_font
            cell.alignment = center_align
            cell.border = cell_border
            
    # Auto-fit column widths
    for col in ws.columns:
        max_len = max(len(str(cell.value or '')) for cell in col)
        col_letter = get_column_letter(col[0].column)
        ws.column_dimensions[col_letter].width = max(max_len + 4, 12)

if inv_file is not None and sales_file is not None:
    try:
        inv_df = load_inventory_data(inv_file)
        sales_df = load_sales_data(sales_file)
        
        # Aggregate Sales by Item ID
        sales_summary_agg = sales_df.groupby('Item ID').agg(
            Total_Sales_Revenue=('Total Gross Bill Amount', 'sum'),
            Total_Orders=('Order Id', 'nunique'),
            Total_Quantity_Sold=('Quantity', 'sum')
        ).reset_index()
        
        # Merge Inventory & Sales
        item_merged = pd.merge(inv_df, sales_summary_agg, on='Item ID', how='left').fillna({
            'Total_Sales_Revenue': 0,
            'Total_Orders': 0,
            'Total_Quantity_Sold': 0,
            'Total Sellable': 0,
            'Warehouse': 0,
            'Darkstore': 0
        })

        # --- EXECUTIVE OVERVIEW ---
        st.markdown("---")
        st.subheader("📊 Executive Dashboard Overview")
        m1, m2, m3, m4, m5 = st.columns(5)
        
        total_sellable = int(inv_df['Total Sellable'].sum()) if 'Total Sellable' in inv_df.columns else 0
        total_wh = int(inv_df['Warehouse'].sum()) if 'Warehouse' in inv_df.columns else 0
        total_ds = int(inv_df['Darkstore'].sum()) if 'Darkstore' in inv_df.columns else 0
        total_rev = sales_df['Total Gross Bill Amount'].sum() if 'Total Gross Bill Amount' in sales_df.columns else 0
        total_orders = sales_df['Order Id'].nunique() if 'Order Id' in sales_df.columns else 0
        
        m1.metric("Sellable Inventory", f"{total_sellable:,}")
        m2.metric("Warehouse Quantity", f"{total_wh:,}")
        m3.metric("Darkstore Quantity", f"{total_ds:,}")
        m4.metric("Total Sales Revenue", f"₹{total_rev:,.2f}")
        m5.metric("Total Orders Mapped", f"{total_orders:,}")
        
        st.markdown("---")

        # --- TABS ---
        tab1, tab2 = st.tabs(["📋 Item-Wise Inv & Sales", "🗺️ State-Wise Inventory Matrix"])
        
        with tab1:
            st.subheader("Item-Wise Inventory Distribution & Sales Mapping")
            
            # Key Columns to display
            display_cols = ['Item ID', 'Item Name', 'Brand Name', 'UoM', 'Warehouse Facility Name', 
                            'Total Sellable', 'Warehouse', 'Darkstore', 
                            'Total_Quantity_Sold', 'Total_Orders', 'Total_Sales_Revenue']
            
            show_cols = [c for c in display_cols if c in item_merged.columns]
            tab1_df = item_merged[show_cols].copy()
            
            st.dataframe(tab1_df, use_container_width=True)
            
        with tab2:
            st.subheader("State-Wise Inventory & Sales Matrix")
            
            state_col = "Supply State" if "Supply State" in sales_df.columns else "Customer State" if "Customer State" in sales_df.columns else None
            
            if state_col:
                all_states = sorted(sales_df[state_col].dropna().unique().tolist())
                selected_state = st.selectbox("Select State Filter", ["All States"] + all_states)
                
                filtered_sales = sales_df if selected_state == "All States" else sales_df[sales_df[state_col] == selected_state]
                
                state_agg = filtered_sales.groupby('Item ID').agg(
                    State_Gross_Sales=('Total Gross Bill Amount', 'sum'),
                    State_Orders=('Order Id', 'nunique'),
                    State_Qty_Sold=('Quantity', 'sum')
                ).reset_index()
                
                state_merged = pd.merge(inv_df, state_agg, on='Item ID', how='left').fillna(0)
                
                state_display_cols = ['Item ID', 'Item Name', 'Brand Name', 'UoM', 
                                      'Total Sellable', 'Warehouse', 'Darkstore', 
                                      'State_Qty_Sold', 'State_Orders', 'State_Gross_Sales']
                
                tab2_df = state_merged[[c for c in state_display_cols if c in state_merged.columns]].copy()
                st.dataframe(tab2_df, use_container_width=True)
            else:
                st.warning("No State column detected in the uploaded sales report.")
                tab2_df = pd.DataFrame()

        # --- EXCEL DOWNLOAD SECTION (Multi-Tab Formatted Excel File) ---
        st.markdown("### 📥 Export Formatted Master Excel Report")
        st.markdown("Download a fully formatted Excel workbook with colored headers, cell borders, centered alignments, and separate analysis tabs.")
        
        excel_buffer = io.BytesIO()
        with pd.ExcelWriter(excel_buffer, engine='openpyxl') as writer:
            # Overview Sheet
            overview_data = pd.DataFrame({
                "KPI Metric": ["Total Sellable Inventory", "Total Warehouse Quantity", "Total Darkstore Quantity", "Total Revenue", "Total Orders Mapped"],
                "Value": [total_sellable, total_wh, total_ds, total_rev, total_orders]
            })
            overview_data.to_excel(writer, sheet_name="Overview KPIs", index=False)
            
            # Tab 1 Sheet
            tab1_df.to_excel(writer, sheet_name="Item_Inv_Distribution", index=False)
            
            # Tab 2 Sheet
            if not tab2_df.empty:
                tab2_df.to_excel(writer, sheet_name="State_Wise_Inventory", index=False)
                
            wb = writer.book
            
            # Apply styling to all sheets
            for sheet_name in wb.sheetnames:
                ws = wb[sheet_name]
                style_excel_sheet(ws, header_bg_color="1F4E78")

        excel_buffer.seek(0)
        
        st.download_button(
            label="📥 Download Master Excel Report (.xlsx)",
            data=excel_buffer,
            file_name="Quick_Commerce_Inventory_Sales_Master_Report.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

    except Exception as e:
        st.error(f"Error processing files: {e}")
else:
    st.info("👈 Please upload both **InventoryData** and **Sales Summary** Excel files in the sidebar to begin.")
