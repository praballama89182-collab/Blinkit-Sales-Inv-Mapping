import io
import openpyxl
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
import pandas as pd
import streamlit as st

st.set_page_config(
    page_title="Quick-Commerce Executive Inventory & Sales Analytics",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded",
)

# Custom Styling & Palette
st.markdown(
    """
    <style>
    .main { background-color: #F8F9FA; }
    .stMetric {
        background-color: #FFFFFF;
        padding: 12px 18px;
        border-radius: 8px;
        box-shadow: 0 2px 4px rgba(0,0,0,0.05);
        border-top: 4px solid #1F4E78;
        margin-bottom: 10px;
    }
    </style>
""",
    unsafe_allow_html=True,
)

st.title("📊 Quick-Commerce Executive Inventory & Sales Dashboard")
st.markdown(
    "Streamlined analysis of SKU variation inventory, darkstore/warehouse"
    " buffers, state coverage, and sales performance."
)

# Sidebar Data Uploads
st.sidebar.header("📁 Data Upload Center")
inv_file = st.sidebar.file_uploader(
    "Upload Inventory Report (.xlsx)", type=["xlsx", "xls"]
)
sales_file = st.sidebar.file_uploader(
    "Upload Sales Summary Report (.xlsx)", type=["xlsx", "xls"]
)


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
  df = df.loc[:, ~df.columns.str.contains("^Unnamed")]

  col_map = {}
  for col in df.columns:
    cl = str(col).strip().lower()
    if "item id" in cl:
      col_map[col] = "Item ID"
    elif "item name" in cl:
      col_map[col] = "Item Name"
    elif "brand" in cl:
      col_map[col] = "Brand Name"
    elif "upc" in cl:
      col_map[col] = "UPC"
    elif "uom" in cl:
      col_map[col] = "UoM"
    elif "total sellable" in cl:
      col_map[col] = "Total Sellable"
    elif cl == "warehouse":
      col_map[col] = "Warehouse"
    elif "darkstore" in cl:
      col_map[col] = "Darkstore"
    elif "last 30 days" in cl:
      col_map[col] = "Last 30 Days Sales"
  df = df.rename(columns=col_map)

  # CRITICAL: Aggregate inventory across multiple warehouse facilities per Item ID
  df_agg = (
      df.groupby(["Item ID", "Brand Name", "Item Name", "UoM"])
      .agg(
          Total_Sellable=("Total Sellable", "sum"),
          Warehouse=("Warehouse", "sum"),
          Darkstore=("Darkstore", "sum"),
          Last_30_Days_Sales=("Last 30 Days Sales", "sum"),
      )
      .reset_index()
  )
  return df_agg


@st.cache_data
def load_sales(file):
  df = pd.read_excel(file, sheet_name=0)
  if "Item Id" in df.columns:
    df = df.rename(columns={"Item Id": "Item ID"})
  return df


def style_excel(ws, header_color="1F4E78"):
  header_fill = PatternFill(
      start_color=header_color, end_color=header_color, fill_type="solid"
  )
  header_font = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
  cell_font = Font(name="Calibri", size=10)

  center_align = Alignment(
      horizontal="center", vertical="center", wrap_text=True
  )
  left_align = Alignment(horizontal="left", vertical="center", wrap_text=True)

  thin_side = Side(style="thin", color="D3D3D3")
  cell_border = Border(
      left=thin_side, right=thin_side, top=thin_side, bottom=thin_side
  )

  for row in ws.iter_rows(
      min_row=1, max_row=ws.max_row, min_col=1, max_col=ws.max_column
  ):
    for cell in row:
      if cell.row == 1 or cell.row == 2:
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = center_align
      else:
        cell.font = cell_font
        if cell.column == 1:
          cell.alignment = left_align
        else:
          cell.alignment = center_align
      cell.border = cell_border

  for col in ws.columns:
    max_l = max(len(str(cell.value or "")) for cell in col)
    col_letter = get_column_letter(col[0].column)
    ws.column_dimensions[col_letter].width = max(max_l + 3, 14)


if inv_file is not None and sales_file is not None:
  try:
    inv_df = load_inventory(inv_file)
    sales_df = load_sales(sales_file)

    state_col = (
        "Supply State"
        if "Supply State" in sales_df.columns
        else "Customer State"
        if "Customer State" in sales_df.columns
        else None
    )

    # --- EXECUTIVE OVERVIEW KPIs (2 Rows) ---
    tot_sellable = int(inv_df["Total_Sellable"].sum())
    tot_wh = int(inv_df["Warehouse"].sum())
    tot_ds = int(inv_df["Darkstore"].sum())
    tot_rev = (
        sales_df["Total Gross Bill Amount"].sum()
        if "Total Gross Bill Amount" in sales_df.columns
        else 0
    )
    tot_orders = (
        sales_df["Order Id"].nunique() if "Order Id" in sales_df.columns else 0
    )
    tot_states = sales_df[state_col].nunique() if state_col else 0

    st.markdown("---")
    st.subheader("📌 Executive Dashboard Overview")

    r1_c1, r1_c2, r1_c3 = st.columns(3)
    r1_c1.metric("Total Sellable Inv", f"{tot_sellable:,}")
    r1_c2.metric("Warehouse Quantity", f"{tot_wh:,}")
    r1_c3.metric("Darkstore Quantity", f"{tot_ds:,}")

    r2_c1, r2_c2, r2_c3 = st.columns(3)
    r2_c1.metric("Total Gross Revenue", f"₹{tot_rev:,.2f}")
    r2_c2.metric("Total Mapped Orders", f"{tot_orders:,}")
    r2_c3.metric("Active Selling States", f"{tot_states} States")
    st.markdown("---")

    # --- PREPARE MERGED DATASETS ---
    sales_orders_item = (
        sales_df.groupby("Item ID")
        .agg(
            Total_Sales_Orders=("Order Id", "nunique"),
            Total_Qty_Sold=("Quantity", "sum"),
            Total_Revenue=("Total Gross Bill Amount", "sum"),
        )
        .reset_index()
    )

    merged_item = pd.merge(
        inv_df, sales_orders_item, on="Item ID", how="left"
    ).fillna(0)

    if state_col:
      item_state_counts = (
          sales_df.groupby("Item ID")[state_col]
          .nunique()
          .reset_index()
          .rename(columns={state_col: "States_Catered_Count"})
      )
      merged_item = pd.merge(
          merged_item, item_state_counts, on="Item ID", how="left"
      ).fillna({"States_Catered_Count": 0})
    else:
      merged_item["States_Catered_Count"] = 0

    # --- TAB STRUCTURE ---
    tab_ov, tab_ps, tab1, tab2, tab_prod = st.tabs([
        "📊 Unique Item Summary",
        "🔍 Product State Coverage & Breakdown",
        "📦 SKU Variation & Sales Mapping",
        "🗺️ State-Wise Inventory & Sales Matrix",
        "📦 Product-Wise Expandable State Metrics",
    ])

    # --- TAB 1: Unique Item Summary ---
    with tab_ov:
      st.subheader("Unique Item Overview (Metrics in Columns)")
      overview_df = (
          merged_item.groupby("Item Name")
          .agg(
              States_Catered=("States_Catered_Count", "max"),
              Total_Sellable=("Total_Sellable", "sum"),
              Warehouse_Quantity=("Warehouse", "sum"),
              Darkstore_Quantity=("Darkstore", "sum"),
              Last_30_Days_Inv_Sales=("Last_30_Days_Sales", "sum"),
              Mapped_Sales_Orders=("Total_Sales_Orders", "sum"),
              Total_Qty_Sold=("Total_Qty_Sold", "sum"),
              Total_Gross_Revenue=("Total_Revenue", "sum"),
          )
          .reset_index()
      )
      st.dataframe(overview_df, use_container_width=True, hide_index=True)

    # --- TAB 2: Product State Coverage & Drilldown ---
    with tab_ps:
      st.subheader("Item State Catering & Interactive Breakdown")
      available_items = sorted(
          merged_item["Item Name"].dropna().unique().tolist()
      )
      selected_product = st.selectbox(
          "🎯 Select Product Name to Filter & Expand State Breakdown:",
          ["All Products"] + available_items,
      )

      if state_col:
        ps_sales = (
            sales_df.groupby(["Item ID", state_col])
            .agg(
                State_Units_Sold=("Quantity", "sum"),
                State_Orders=("Order Id", "nunique"),
                State_Gross_Revenue=("Total Gross Bill Amount", "sum"),
            )
            .reset_index()
        )

        ps_full = (
            pd.merge(
                merged_item[[
                    "Item ID",
                    "Brand Name",
                    "Item Name",
                    "UoM",
                    "Total_Sellable",
                    "Warehouse",
                    "Darkstore",
                ]],
                ps_sales,
                on="Item ID",
                how="left",
            )
            .rename(
                columns={
                    state_col: "Catered State",
                    "Total_Sellable": "Total Sellable",
                }
            )
            .fillna(0)
        )

        product_state_summary_df = (
            merged_item.groupby(["Brand Name", "Item Name", "UoM"])
            .agg(
                Total_States_Catered=("States_Catered_Count", "max"),
                Total_Sellable=("Total_Sellable", "sum"),
                Warehouse_Stock=("Warehouse", "sum"),
                Darkstore_Stock=("Darkstore", "sum"),
                Total_30D_Orders=("Total_Sales_Orders", "sum"),
                Total_Units_Sold=("Total_Qty_Sold", "sum"),
                Total_Revenue=("Total_Revenue", "sum"),
            )
            .reset_index()
        )

        if selected_product != "All Products":
          filtered_summary = product_state_summary_df[
              product_state_summary_df["Item Name"] == selected_product
          ]
          filtered_breakdown = ps_full[
              ps_full["Item Name"] == selected_product
          ]
          st.markdown(
              f"#### 📌 Selected Product Overview: `{selected_product}`"
          )
          st.dataframe(
              filtered_summary, use_container_width=True, hide_index=True
          )
          st.markdown(
              f"#### 🌐 States Catered Breakdown for `{selected_product}`"
          )
          st.dataframe(
              filtered_breakdown[[
                  "Item Name",
                  "Catered State",
                  "UoM",
                  "Total Sellable",
                  "Warehouse",
                  "Darkstore",
                  "State_Units_Sold",
                  "State_Orders",
                  "State_Gross_Revenue",
              ]],
              use_container_width=True,
              hide_index=True,
          )
        else:
          st.markdown("#### 📋 All Products State Coverage Summary")
          st.dataframe(
              product_state_summary_df,
              use_container_width=True,
              hide_index=True,
          )
          st.markdown("#### 🌐 Complete Product-State Breakdown Table")
          st.dataframe(
              ps_full[[
                  "Item Name",
                  "Catered State",
                  "UoM",
                  "Total Sellable",
                  "Warehouse",
                  "Darkstore",
                  "State_Units_Sold",
                  "State_Orders",
                  "State_Gross_Revenue",
              ]],
              use_container_width=True,
              hide_index=True,
          )
      else:
        st.warning("No state column available.")
        ps_full = pd.DataFrame()

    # --- TAB 3: SKU Variation & Sales Mapping ---
    with tab1:
      st.subheader("SKU Variation Inventory & Sales Mapping")
      tab1_df = merged_item.rename(
          columns={
              "UoM": "Variation (Pack Size)",
              "Total_Sellable": "Total Sellable",
              "States_Catered_Count": "States Catered",
              "Last_30_Days_Sales": "Last 30D Inv Movement",
              "Total_Sales_Orders": "30D Mapped Orders",
              "Total_Qty_Sold": "30D Units Sold",
              "Total_Revenue": "30D Revenue (Rs)",
          }
      )
      st.dataframe(tab1_df, use_container_width=True, hide_index=True)

    # --- TAB 4: State-Wise Matrix ---
    with tab2:
      st.subheader("State-Wise Matrix (States Above Metrics)")
      if state_col:
        sales_state_pivot = sales_df.pivot_table(
            index="Item ID",
            columns=state_col,
            values=["Quantity", "Order Id", "Total Gross Bill Amount"],
            aggfunc={
                "Quantity": "sum",
                "Order Id": "nunique",
                "Total Gross Bill Amount": "sum",
            },
            fill_value=0,
        )
        sales_state_pivot.columns = sales_state_pivot.columns.set_levels(
            ["Orders", "Qty Sold", "Revenue (Rs)"], level=0
        )
        flat_cols = [
            f"{state} - {metric}" for metric, state in sales_state_pivot.columns
        ]
        sales_state_pivot.columns = flat_cols
        sales_state_pivot = sales_state_pivot.reset_index()

        state_matrix_df = (
            pd.merge(
                inv_df[[
                    "Item ID",
                    "Brand Name",
                    "Item Name",
                    "UoM",
                    "Total_Sellable",
                    "Warehouse",
                    "Darkstore",
                ]],
                sales_state_pivot,
                on="Item ID",
                how="left",
            )
            .fillna(0)
            .rename(columns={"Total_Sellable": "Total Sellable"})
            .drop(columns=["Item ID"])
        )
        st.dataframe(
            state_matrix_df,
            use_container_width=True,
            hide_index=True,
            column_config={
                "Item Name": st.column_config.Column(pinned=True),
                "Brand Name": st.column_config.Column(pinned=True),
            },
        )
      else:
        state_matrix_df = pd.DataFrame()

    # --- TAB 5: Product-Wise Expandable State Metrics ---
    with tab_prod:
      st.subheader("📦 Product-Wise Metrics & State Breakdown")
      prod_summary = (
          merged_item.groupby("Item Name")
          .agg(
              States_Catered=("States_Catered_Count", "max"),
              Total_Sellable=("Total_Sellable", "sum"),
              Warehouse_Stock=("Warehouse", "sum"),
              Darkstore_Stock=("Darkstore", "sum"),
              Total_Orders=("Total_Sales_Orders", "sum"),
              Total_Units_Sold=("Total_Qty_Sold", "sum"),
              Total_Revenue=("Total_Revenue", "sum"),
          )
          .reset_index()
      )
      st.markdown("#### 📋 Product High-Level Summary")
      st.dataframe(prod_summary, use_container_width=True, hide_index=True)

      st.markdown("#### 🔽 Interactive Expandable Product State Metrics")
      if state_col:
        ps_df = (
            sales_df.groupby(["Item ID", state_col])
            .agg(
                State_Units_Sold=("Quantity", "sum"),
                State_Orders=("Order Id", "nunique"),
                State_Revenue=("Total Gross Bill Amount", "sum"),
            )
            .reset_index()
            .rename(columns={state_col: "State"})
        )
        ps_merged_full = (
            pd.merge(
                merged_item[[
                    "Item ID",
                    "Item Name",
                    "UoM",
                    "Total_Sellable",
                    "Warehouse",
                    "Darkstore",
                ]],
                ps_df,
                on="Item ID",
                how="left",
            )
            .rename(columns={"Total_Sellable": "Total Sellable"})
            .fillna(0)
        )

        for prod_name in sorted(merged_item["Item Name"].unique()):
          prod_sub_df = ps_merged_full[
              ps_merged_full["Item Name"] == prod_name
          ]
          tot_p_sellable = (
              int(prod_sub_df["Total Sellable"].iloc[0])
              if not prod_sub_df.empty
              else 0
          )
          tot_p_wh = (
              int(prod_sub_df["Warehouse"].iloc[0])
              if not prod_sub_df.empty
              else 0
          )
          tot_p_ds = (
              int(prod_sub_df["Darkstore"].iloc[0])
              if not prod_sub_df.empty
              else 0
          )
          tot_p_rev = prod_sub_df["State_Revenue"].sum()
          n_states = prod_sub_df["State"].nunique()

          with st.expander(
              f"📦 **{prod_name}** | States: {n_states} | Sellable:"
              f" {tot_p_sellable} | WH: {tot_p_wh} | DS: {tot_p_ds} | Rev:"
              f" ₹{tot_p_rev:,.2f}"
          ):
            display_state_sub = prod_sub_df[[
                "State",
                "UoM",
                "Total Sellable",
                "Warehouse",
                "Darkstore",
                "State_Units_Sold",
                "State_Orders",
                "State_Revenue",
            ]].rename(
                columns={
                    "UoM": "Variation (Pack)",
                    "State_Units_Sold": "Units Sold",
                    "State_Orders": "Orders",
                    "State_Revenue": "Revenue (Rs)",
                }
            )
            st.dataframe(
                display_state_sub, use_container_width=True, hide_index=True
            )

    # --- MASTER EXPORT SECTION ---
    st.markdown("---")
    st.markdown("### 📥 Download Formatted Master Excel Report")
    excel_buf = io.BytesIO()
    with pd.ExcelWriter(excel_buf, engine="openpyxl") as writer:
      overview_df.to_excel(
          writer, sheet_name="Unique_Item_Overview", index=False
      )
      if "ps_full" in locals() and not ps_full.empty:
        ps_full.drop(columns=["Item ID"], errors="ignore").to_excel(
            writer, sheet_name="Product_State_Breakdown", index=False
        )
      tab1_df.drop(columns=["Item ID"], errors="ignore").to_excel(
          writer, sheet_name="SKU_Variation_Sales", index=False
      )
      if not state_matrix_df.empty:
        state_matrix_df.to_excel(
            writer, sheet_name="State_Wise_Matrix", index=False
        )
      prod_summary.to_excel(
          writer, sheet_name="Product_Wise_Summary", index=False
      )

      wb = writer.book
      for sheet in wb.sheetnames:
        style_excel(wb[sheet])

    excel_buf.seek(0)
    st.download_button(
        label="📥 Export Clean Multi-Tab Excel Workbook (.xlsx)",
        data=excel_buf,
        file_name="Quick_Commerce_Inventory_Sales_Master.xlsx",
        mime=(
            "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        ),
    )

  except Exception as e:
    st.error(f"Error executing dashboard pipeline: {e}")
else:
  st.info(
      "👈 Upload **Inventory Report** and **Sales Summary** Excel files in the"
      " sidebar to generate the interactive dashboard."
  )
