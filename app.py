import io
import pandas as pd
import streamlit as st
import openpyxl
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
from openpyxl.utils import get_column_letter

st.set_page_config(
    page_title="Quick-Commerce Executive Inventory & Sales Analytics",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded"
)

st.markdown("""
    <style>
    .main { background-color: #F8F9FA; }
    .stMetric {
        background-color: #FFFFFF;
        padding: 15px;
        border-radius: 8px;
        box-shadow: 0 2px 4px rgba(0,0,0,0.05);
        border-top: 4px solid #1F4E78;
    }
    .css-1544g2n { padding: 1rem 1rem; }
    </style>
""", unsafe_allow_html=True)

st.title("📊 Quick-Commerce Executive Inventory & Sales Dashboard")
st.markdown("Streamlined analysis of SKU variation inventory, darkstore/warehouse buffers, and state-level quick-commerce performance.")

st.sidebar.header("📁 Data Upload Center")
inv_file = st.sidebar.file_uploader("Upload Inventory Report (.xlsx)", type=["xlsx", "xls"])
sales_file = st.sidebar.file_uploader("Upload Sales Summary Report (.xlsx)", type=["xlsx", "xls"])

@st.cache_data
def load_inventory(file):
    df_raw = pd.read_excel(file, sheet_name=0, header=None)
    header_idx = 2
    for idx, row in df_raw.iterrows():
        row_vals = [str(v).strip().lower() for v in row.values if pd.notna(v)]
        if "item id" in row_vals:
            header_idx = idx
            break
    df = pd.read_excel(file, sheet_name=0, header=header_idx)
    df = df.loc[:, ~df.columns.str.contains('^Unnamed')]
    
    col_map = {}
    for col in df.columns:
        cl = str(col).strip().lower()
        if "item id" in cl: col_map[col] = "Item ID"
        elif "item name" in cl: col_map[col] = "Item Name"
        elif "brand" in cl: col_map[col] = "Brand Name"
        elif "upc" in cl: col_map[col] = "UPC"
        elif "uom" in cl: col_map[col] = "UoM"
        elif "total sellable" in cl: col_map[col] = "Total Sellable"
        elif cl == "warehouse": col_map[col] = "Warehouse"
        elif "darkstore" in cl: col_map[col] = "Darkstore"
        elif "last 30 days" in cl: col_map[col] = "Last 30 Days Sales"
    return df.rename(columns=col_map)

@st.cache_data
def load_sales(file):
    df = pd.read_excel(file, sheet_name=0)
    if "Item Id" in df.columns:
        df = df.rename(columns={"Item Id": "Item ID"})
    return df

def style_excel(ws, header_color="1F4E78"):
    header_fill = PatternFill(start_color=header_color, end_color=header_color, fill_type="solid")
    header_font = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
    cell_font = Font(name="Calibri", size=10)
    center = Alignment(horizontal="center", vertical="center", wrap_text=True)
    thin = Side(style='thin', color='D3D3D3')
    border = Border(left=thin, right=thin, top=thin, bottom=thin)
    
    for row in ws.iter_rows(min_row=1, max_row=ws.max_row, min_col=1, max_col=ws.max_column):
        for cell in row:
            if cell.row == 1 or cell.row == 2:
                cell.fill = header_fill
                cell.font = header_font
            else:
                cell.font = cell_font
            cell.alignment = center
            cell.border = border

    for col in ws.columns:
        max_l = max(len(str(cell.value or '')) for cell in col)
        col_letter = get_column_letter(col[0].column)
        ws.column_dimensions[col_letter].width = max(max_l + 3, 12)

if inv_file is not None and sales_file is not None:
    try:
        inv_df = load_inventory(inv_file)
        sales_df = load_sales(sales_file)
        
        # --- EXECUTIVE OVERVIEW KPIs ---
        tot_sellable = int(inv_df['Total Sellable'].sum()) if 'Total Sellable' in inv_df.columns else 0
        tot_wh = int(inv_df['Warehouse'].sum()) if 'Warehouse' in inv_df.columns else 0
        tot_ds = int(inv_df['Darkstore'].sum()) if 'Darkstore' in inv_df.columns else 0
        tot_rev = sales_df['Total Gross Bill Amount'].sum() if 'Total Gross Bill Amount' in sales_df.columns else 0
        tot_orders = sales_df['Order Id'].nunique() if 'Order Id' in sales_df.columns else 0
        
        st.markdown("---")
        st.subheader("📌 Executive Dashboard Overview")
        k1, k2, k3, k4, k5 = st.columns(5)
        k1.metric("Total Sellable Inv", f"{tot_sellable:,}")
        k2.metric("Warehouse Quantity", f"{tot_wh:,}")
        k3.metric("Darkstore Quantity", f"{tot_ds:,}")
        k4.metric("Total Gross Revenue", f"₹{tot_rev:,.2f}")
        k5.metric("Total Mapped Orders", f"{tot_orders:,}")
        st.markdown("---")
        
        # --- TAB STRUCTURE ---
        tab_ov, tab1, tab2, tab3 = st.tabs([
            "📊 Unique Item Overview",
            "📦 SKU Variation & Sales Mapping",
            "🗺️ State-Wise Inventory & Sales Matrix",
            "🏷️ Brand-Wise Summary & Velocity"
        ])
        
        # --- OVERVIEW TAB: Unique Item Name Summary ---
        with tab_ov:
            st.subheader("Unique Item Summary (Metrics in Columns)")
            st.markdown("Clean unique item-level aggregation showing Total Sellable, Warehouse, Darkstore, and 30-Day Sales Orders.")
            
            # Map 30D Sales orders from Sales summary
            sales_orders_item = sales_df.groupby('Item ID').agg(
                Total_Sales_Orders=('Order Id', 'nunique'),
                Total_Qty_Sold=('Quantity', 'sum'),
                Total_Revenue=('Total Gross Bill Amount', 'sum')
            ).reset_index()
            
            merged_item = pd.merge(inv_df, sales_orders_item, on='Item ID', how='left').fillna(0)
            
            overview_df = merged_item.groupby('Item Name').agg(
                Total_Sellable=('Total Sellable', 'sum'),
                Warehouse_Quantity=('Warehouse', 'sum'),
                Darkstore_Quantity=('Darkstore', 'sum'),
                Last_30_Days_Inv_Sales=('Last 30 Days Sales', 'sum'),
                Mapped_Sales_Orders=('Total_Sales_Orders', 'sum'),
                Total_Qty_Sold=('Total_Qty_Sold', 'sum'),
                Total_Gross_Revenue=('Total_Revenue', 'sum')
            ).reset_index()
            
            st.dataframe(overview_df, use_container_width=True, hide_index=True)
            
        # --- TAB 1: Variation (Brand, Item, Pack/UoM) ---
        with tab1:
            st.subheader("SKU Variation Inventory & Sales Mapping")
            st.markdown("Granular breakdown formatted as **Brand Name | Item Name | Variation (UoM)** with respective stock levels and sales performance.")
            
            tab1_cols = ['Brand Name', 'Item Name', 'UoM', 'Total Sellable', 'Warehouse', 'Darkstore', 'Last 30 Days Sales', 'Total_Sales_Orders', 'Total_Qty_Sold', 'Total_Revenue']
            existing_tab1_cols = [c for c in tab1_cols if c in merged_item.columns]
            
            tab1_df = merged_item[existing_tab1_cols].rename(columns={
                'UoM': 'Variation (Pack Size)',
                'Last 30 Days Sales': 'Last 30D Inv Movement',
                'Total_Sales_Orders': '30D Mapped Orders',
                'Total_Qty_Sold': '30D Units Sold',
                'Total_Revenue': '30D Revenue (Rs)'
            })
            
            st.dataframe(tab1_df, use_container_width=True, hide_index=True)

        # --- TAB 2: State-Wise Matrix with Right-Scrollable Columns & Frozen Item Name ---
        with tab2:
            st.subheader("State-Wise Matrix (States Above Metrics)")
            st.markdown("Scroll right to view state performance. The **Item Name** stays fixed on the left.")
            
            state_col = 'Supply State' if 'Supply State' in sales_df.columns else 'Customer State'
            
            # Pivot sales by Item ID & State
            sales_state_pivot = sales_df.pivot_table(
                index='Item ID',
                columns=state_col,
                values=['Quantity', 'Order Id', 'Total Gross Bill Amount'],
                aggfunc={'Quantity': 'sum', 'Order Id': 'nunique', 'Total Gross Bill Amount': 'sum'},
                fill_value=0
            )
            
            # Rename level metrics
            sales_state_pivot.columns = sales_state_pivot.columns.set_levels(
                ['Orders', 'Qty Sold', 'Revenue (Rs)'], level=0
            )
            
            # Flatten MultiIndex columns with State on Top
            # Form: [State] [Metric]
            flat_cols = [f"{state} - {metric}" for metric, state in sales_state_pivot.columns]
            sales_state_pivot.columns = flat_cols
            sales_state_pivot = sales_state_pivot.reset_index()
            
            state_matrix_df = pd.merge(
                inv_df[['Item ID', 'Brand Name', 'Item Name', 'UoM', 'Total Sellable', 'Warehouse', 'Darkstore']],
                sales_state_pivot,
                on='Item ID',
                how='left'
            ).fillna(0)
            
            # Remove Item ID for clean viewing
            state_matrix_df = state_matrix_df.drop(columns=['Item ID'])
            
            st.dataframe(
                state_matrix_df,
                use_container_width=True,
                hide_index=True,
                column_config={
                    "Item Name": st.column_config.Column(pinned=True),
                    "Brand Name": st.column_config.Column(pinned=True)
                }
            )

        # --- TAB 3: Brand-Wise Performance & Inventory Health ---
        with tab3:
            st.subheader("Brand-Wise Inventory Distribution & Channel Velocity")
            st.markdown("Strategic brand-level aggregation to compare total darkstore vs warehouse distribution and revenue contribution.")
            
            brand_summary = merged_item.groupby('Brand Name').agg(
                Total_SKUs=('Item ID', 'nunique'),
                Total_Sellable_Stock=('Total Sellable', 'sum'),
                Warehouse_Stock=('Warehouse', 'sum'),
                Darkstore_Stock=('Darkstore', 'sum'),
                Total_30D_Orders=('Total_Sales_Orders', 'sum'),
                Total_30D_Units=('Total_Qty_Sold', 'sum'),
                Total_30D_Revenue=('Total_Revenue', 'sum')
            ).reset_index()
            
            st.dataframe(brand_summary, use_container_width=True, hide_index=True)

        # --- MASTER EXPORT SECTION ---
        st.markdown("---")
        st.markdown("### 📥 Download Formatted Master Excel Report")
        st.markdown("Export all cleaned analysis tabs with dark navy headers, cell borders, centered alignments, and auto-fitted columns.")
        
        excel_buf = io.BytesIO()
        with pd.ExcelWriter(excel_buf, engine='openpyxl') as writer:
            overview_df.to_excel(writer, sheet_name="Unique_Item_Overview", index=False)
            tab1_df.to_excel(writer, sheet_name="SKU_Variation_Sales", index=False)
            state_matrix_df.to_excel(writer, sheet_name="State_Wise_Matrix", index=False)
            brand_summary.to_excel(writer, sheet_name="Brand_Summary", index=False)
            
            wb = writer.book
            for sheet in wb.sheetnames:
                style_excel(wb[sheet])
                
        excel_buf.seek(0)
        st.download_button(
            label="📥 Export Clean Multi-Tab Excel Workbook (.xlsx)",
            data=excel_buf,
            file_name="Quick_Commerce_Inventory_Sales_Master.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
        
    except Exception as e:
        st.error(f"Error executing dashboard pipeline: {e}")
else:
    st.info("👈 Upload **Inventory Report** and **Sales Summary** Excel files in the sidebar to generate the interactive dashboard.")
