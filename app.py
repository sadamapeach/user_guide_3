import streamlit as st
import pandas as pd
import numpy as np
import time
import re
from io import BytesIO

def format_rupiah(x):
    if pd.isna(x):
        return ""
    # pastikan bisa diubah ke float
    try:
        x = float(x)
    except:
        return x  # biarin apa adanya kalau bukan angka

    # kalau tidak punya desimal (misal 7000.0), tampilkan tanpa ,00
    if x.is_integer():
        formatted = f"{int(x):,}".replace(",", ".")
    else:
        formatted = f"{x:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
        # hapus ,00 kalau desimalnya 0 semua (misal 7000,00 → 7000)
        if formatted.endswith(",00"):
            formatted = formatted[:-3]
    return formatted

def highlight_total(row):
    # Cek apakah ada kolom yang berisi "TOTAL" (case-insensitive)
    if any(str(x).strip().upper() == "TOTAL" for x in row):
        return ["font-weight: bold; background-color: #D9EAD3; color: #1A5E20;"] * len(row)
    else:
        return [""] * len(row)
    
# Highlight total per year
def highlight_total_per_year(row):
    if str(row["SCOPE"]).strip().upper() == "TOTAL" and pd.notna(row["YEAR"]):
        return ["font-weight: bold; background-color: #FFEB9C; color: #9C6500;"] * len(row)
    else:
        return [""] * len(row)

# Highlight vendor total
def highlight_vendor_total(row):
    if str(row["YEAR"]).strip().upper() == "TOTAL":
        return ["font-weight: bold; background-color: #C6EFCE; color: #006100;"] * len(row)
    else:
        return [""] * len(row)
    
def highlight_1st_2nd_vendor(row, columns):
    styles = [""] * len(columns)
    first_vendor = row.get("1st Vendor")
    second_vendor = row.get("2nd Vendor")

    for i, col in enumerate(columns):
        if col == first_vendor:
            # styles[i] = "background-color: #f8c8dc; color: #7a1f47;"
            styles[i] = "background-color: #C6EFCE; color: #006100;"
        elif col == second_vendor:
            # styles[i] = "background-color: #d7c6f3; color: #402e72;"
            styles[i] = "background-color: #FFEB9C; color: #9C6500;"
    return styles

st.subheader("🧑‍🏫 User Guide: TCO Comparison by Year + Region")
st.markdown(
    ":red-badge[Indosat] :orange-badge[Ooredoo] :green-badge[Hutchison]"
)
st.caption("INSPIRE 2025 | Oktaviana Sadama Nur Azizah")

# Divider custom
st.markdown(
    """
    <hr style="margin-top:-5px; margin-bottom:10px; border: none; height: 2px; background-color: #ddd;">
    """,
    unsafe_allow_html=True
)

st.markdown(
    """
    <div style="
        display: flex;
        align-items: center;
        height: 65px;
        margin-bottom: 10px;
    ">
        <div style="text-align: justify; font-size: 15px;">
            <span style="color: #FF1BF1; font-weight: 800;">
            TCO Comparison by Year + Region</span>
            combines yearly and regional TCO analysis to provide a more comprehensive view of 
            vendor pricing across time and location.
        </div>
    </div>
    """,
    unsafe_allow_html=True
)

st.markdown("#### Input Structure")

st.markdown(
    """
        <div style="text-align: justify; font-size: 15px; margin-bottom: 20px">
            The input file required for this menu should be a <span style="color: #FF69B4; font-weight: 500;">
            single file containing multiple sheets</span>, in eather <span style="background:#C6EFCE; 
            padding:1px 4px; border-radius:6px; font-weight:600; font-size: 0.75rem; color: black">.xlsx</span> 
            or <span style="background:#FFEB9C; padding:2px 4px; border-radius:6px; font-weight:600; 
            font-size: 0.75rem; color: black">.xls</span> format. Each sheet represents a vendor name, with the 
            table structure in each sheet as follows:
        </div>
    """,
    unsafe_allow_html=True
)

# Dataframe
columns = ["Year", "Scope", "Desc", "Region 1", "Region 2", "Region 3", "Region 4", "Region 5"]
df = pd.DataFrame([[""] * len(columns) for _ in range(3)], columns=columns)

st.dataframe(df, hide_index=True)

# Buat DataFrame 1 row
st.markdown("""
<table style="width: 100%; border-collapse: collapse; table-layout: fixed; font-size: 15px;">
    <tr>
        <td style="border: 1px solid gray; width: 15%;">Vendor A</td>
        <td style="border: 1px solid gray; width: 15%;">Vendor B</td>
        <td style="border: 1px solid gray; width: 15%;">Vendor C</td>
        <td style="border: 1px solid gray; font-style: italic; color: #26BDAD">multiple sheets</td>
    </tr>
</table>
""", unsafe_allow_html=True)

st.markdown("###### Description:")
st.markdown(
    """
    <div style="font-size:15px;">
        <ul>
            <li>
                <span style="display:inline-block; width:100px;">Year</span>: must be on the left
            </li>
            <li>
                <span style="display:inline-block; width:100px;">Scope & Desc</span>: non-numeric columns
            </li>
            <li>
                <span style="display:inline-block; width:100px;">Region 1 to 5</span>: numeric columns
            </li>
        </ul>
    </div>
    """,
    unsafe_allow_html=True
)

st.markdown(
    """
        <div style="text-align: justify; font-size: 15px; margin-bottom: 20px">
            The system accommodates a <span style="font-weight: bold;">dynamic table</span>,
            allowing users to enter any number of non-numeric and numeric columns. Users have 
            the freedom to name the columns as they wish. The system logic relies on <span 
            style="font-weight: bold;">column indices</span>, not specific column names.
        </div>
    """,
    unsafe_allow_html=True
)

st.markdown("**:violet-badge[Ensure that each sheet has the same table structure and column names!]**")

st.divider()
st.markdown("#### Constraint")

st.markdown(
    """
        <div style="text-align: justify; font-size: 15px; margin-bottom: 20px; margin-top: -10px">
            To ensure this menu works correctly, users need to follow certain rules regarding
            the dataset structure.
        </div>
    """,
    unsafe_allow_html=True
)

st.markdown("**:red-badge[1. COLUMN ORDER]**")
st.markdown(
    """
        <div style="text-align: justify; font-size: 15px; margin-bottom: 10px; margin-top: -10px">
            When creating tables, it is important to follow the specified column structure. Columns 
            <span style="font-weight: bold;">must</span> be arranged in the following order:
        </div>
    """,
    unsafe_allow_html=True
)

st.markdown(
    """
        <div style="text-align: center; font-size: 15px; margin-bottom: 10px; font-weight: bold">
            Year → Non-Numeric Columns → Numeric Columns
        </div>
    """,
    unsafe_allow_html=True
)

st.markdown(
    """
        <div style="text-align: justify; font-size: 15px; margin-bottom: 20px">
            this order is <span style="color: #FF69B4; font-weight: 700;">strict</span> and 
            <span style="color: #FF69B4; font-weight: 700;">cannot be altered</span>!
        </div>
    """,
    unsafe_allow_html=True
)

st.markdown("**:orange-badge[2. NUMBER COLUMN]**")
st.markdown(
    """
        <div style="text-align: justify; font-size: 15px; margin-bottom: 10px; margin-top:-10px;">
            Please refer the table below:
        </div>
    """,
    unsafe_allow_html=True
)

# DataFrame
columns = ["No", "Year", "Scope", "Desc", "Region 1", "Region 2", "Region 3", "Region 4", "Region 5"]
data = [
    [1] + [""] * (len(columns) - 1),
    [2] + [""] * (len(columns) - 1),
    [3] + [""] * (len(columns) - 1)
]
df = pd.DataFrame(data, columns=columns)

st.dataframe(df, hide_index=True)

st.markdown(
    """
        <div style="text-align: justify; font-size: 15px; margin-bottom: 20px; margin-top: -5px;">
            The table above is an <span style="color: #FF69B4; font-weight: 700;">incorrect example</span>
            and is <span style="color: #FF69B4; font-weight: 700;">not allowed</span> because it contains 
            a <span style="font-weight: bold;">"No"</span> column. The "No" column is prohibited in this
            menu, as it will be treated as a numeric column by the system, which violates the constraint
            described in point 1. Additionally, the <span style="background:#FFCB09; padding:2px 4px; 
            border-radius:6px; font-weight:600; font-size: 0.75rem; color: black">Year</span> column must 
            be placed as the first index (idx[0]).
        </div>
    """,
    unsafe_allow_html=True
)

st.markdown("**:green-badge[3. FLOATING TABLE]**")
st.markdown(
    """
        <div style="text-align: justify; font-size: 15px; margin-bottom: 10px; margin-top:-10px;">
            Floating tables are allowed, meaning tables <span style="color: #FF69B4; font-weight: 700;">
            do not need to start from cell A1</span>. However, ensure
            that the cells above and to the left of the table are empty, as shown in the example below:
        </div>
    """,
    unsafe_allow_html=True
)

# DataFrame
columns = ["", "A", "B", "C", "D", "E", "F", "G"]

# Buat 5 baris kosong
df = pd.DataFrame([[""] * len(columns) for _ in range(7)], columns=columns)

# Isi kolom pertama dengan 1–7
df.iloc[:, 0] = [1, 2, 3, 4, 5, 6, 7]

# Header bagian kedua
df.loc[1, ["B", "C", "D", "E", "F"]] = ["Year", "Scope", "Region 1", "Region 2", "Region 3"]

# Data Software & Hardware
df.loc[2, ["B", "C", "D", "E", "F"]] = ["2025", "AirCon Dismantle", "1.000", "2.000", "3.000"]
df.loc[3, ["B", "C", "D", "E", "F"]] = ["2025", "ACPBD Dismantle", "5.500", "6.500", "7.500"]
df.loc[4, ["B", "C", "D", "E", "F"]] = ["2026", "AirCon Dismantle", "1.200", "1.900", "3.100"]
df.loc[5, ["B", "C", "D", "E", "F"]] = ["2026", "ACPBD Dismantle", "5.300", "6.700", "7.300"]

st.dataframe(df, hide_index=True)

st.markdown(
    """
        <div style="text-align: justify; font-size: 15px; margin-bottom: 20px; margin-top:-10px;">
            To provide additional explanations or notes on the sheet, you can include them using an image or a text box.
        </div>
    """,
    unsafe_allow_html=True
)

st.markdown("**:blue-badge[Unlike the 'TCO Comparison by Year' menu, you are NOT allowed to add a TOTAL column as the last column!]**")

st.divider()

st.markdown("#### What is Displayed?")

# Path file Excel yang sudah ada
file_path = "dummy dataset.xlsx"

# Buka file sebagai binary
with open(file_path, "rb") as f:
    file_data = f.read()

# Markdown teks
st.markdown(
    """
    <div style="text-align: justify; font-size: 15px; margin-bottom: 5px; margin-top: -10px">
        You can try this menu by downloading the dummy dataset using the button below: 
    </div>
    """,
    unsafe_allow_html=True
)

@st.fragment
def release_the_balloons():
    st.balloons()

# Download button untuk file Excel
st.download_button(
    label="Dummy Dataset",
    data=file_data,
    file_name="dummy dataset.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    on_click=release_the_balloons,
    type="primary",
    use_container_width=True,
)

st.markdown(
    """
    <div style="text-align: justify; font-size: 15px; margin-bottom: 20px">
        Based on this dummy dataset, the menu will produce the following results.
    </div>
    """,
    unsafe_allow_html=True
)

st.markdown("**:red-badge[1. MERGE DATA]**")
st.markdown(
    """
        <div style="text-align: justify; font-size: 15px; margin-bottom: 10px; margin-top:-10px;">
            The system will merge the tables from each sheet into a single table and add
            a <span style="background:#FFCB09; padding:2px 4px; border-radius:6px; font-weight:600; 
            font-size: 0.75rem; color: black">TOTAL ROW</span> based on year and scope, as shown below.
        </div>
    """,
    unsafe_allow_html=True
)

# DataFrame
columns = ["VENDOR", "YEAR", "SCOPE", "REGION 1", "REGION 2", "TOTAL"]
data = [
    ["Vendor A", "2025", "Site Survey", 8500, 9000, 17500],
    ["Vendor A", "2025", "DG Dismantle", 45000, 47800, 92800],
    ["Vendor A", "2025", "RAN & Power Supply", 68000, 65500, 133500],
    ["Vendor A", "2025", "TOTAL", 121500, 122300, 243800],

    ["Vendor A", "2026", "Site Survey", 9200, 8800, 18000],
    ["Vendor A", "2026", "DG Dismantle", 46500, 44000, 90500],
    ["Vendor A", "2026", "RAN & Power Supply", 74900, 72400, 147300],
    ["Vendor A", "2026", "TOTAL", 130600, 125200, 255800],

    ["Vendor A", "TOTAL", "", 252100, 247500, 499600],

    ["Vendor B", "2025", "Site Survey", 8200, 8700, 16900],
    ["Vendor B", "2025", "DG Dismantle", 46200, 49000, 95200],
    ["Vendor B", "2025", "RAN & Power Supply", 70500, 67900, 138400],
    ["Vendor B", "2025", "TOTAL", 124900, 125600, 250500],

    ["Vendor B", "2026", "Site Survey", 8700, 8300, 17000],
    ["Vendor B", "2026", "DG Dismantle", 48900, 46800, 95700],
    ["Vendor B", "2026", "RAN & Power Supply", 71900, 69300, 141200],
    ["Vendor B", "2026", "TOTAL", 129500, 124400, 253900],

    ["Vendor B", "TOTAL", "", 254400, 250000, 504400],

    ["Vendor C", "2025", "Site Survey", 8900, 9200, 18100],
    ["Vendor C", "2025", "DG Dismantle", 43500, 46000, 89500],
    ["Vendor C", "2025", "RAN & Power Supply", 73800, 71500, 145300],
    ["Vendor C", "2025", "TOTAL", 126200, 126700, 252900],

    ["Vendor C", "2026", "Site Survey", 9700, 9300, 19000],
    ["Vendor C", "2026", "DG Dismantle", 51300, 49100, 100400],
    ["Vendor C", "2026", "RAN & Power Supply", 69500, 66800, 136300],
    ["Vendor C", "2026", "TOTAL", 130500, 125200, 255700],

    ["Vendor C", "TOTAL", "", 256700, 251900, 508600],
]
df_merge = pd.DataFrame(data, columns=columns)

num_cols = ["REGION 1", "REGION 2", "TOTAL"]
df_merge_styled = (
    df_merge.style
    .format({col: format_rupiah for col in num_cols})
    .apply(highlight_total_per_year, axis=1)
    .apply(highlight_vendor_total, axis=1)
)

st.dataframe(df_merge_styled, hide_index=True)

st.write("")
st.markdown("**:orange-badge[2. COST SUMMARY]**")
st.markdown(
    """
        <div style="text-align: justify; font-size: 15px; margin-bottom: 10px; margin-top:-10px;">
            After merging the data, the system will transpose the region columns into a single column and add a 
            <span style="background: #FF5E5E; padding:1px 4px; border-radius:6px; font-weight:600; font-size: 13px; color: black">PRICE</span>
            column as the final column in the table.
        </div>
    """,
    unsafe_allow_html=True
)

# DataFrame
columns = ["VENDOR", "YEAR", "REGION", "SCOPE", "PRICE"]
data = [
    ["Vendor A", "2025", "REGION 1", "Site Survey", 8500],
    ["Vendor A", "2025", "REGION 1", "DG Dismantle", 45000],
    ["Vendor A", "2025", "REGION 1", "RAN & Power Supply", 68000],
    ["Vendor A", "2025", "REGION 1", "TOTAL", 121500],

    ["Vendor A", "2026", "REGION 1", "Site Survey", 9200],
    ["Vendor A", "2026", "REGION 1", "DG Dismantle", 46500],
    ["Vendor A", "2026", "REGION 1", "RAN & Power Supply", 74900],
    ["Vendor A", "2026", "REGION 1", "TOTAL", 130600],

    ["Vendor A", "TOTAL", "REGION 1", "", 252100],

    ["Vendor B", "2025", "REGION 1", "Site Survey", 8200],
    ["Vendor B", "2025", "REGION 1", "DG Dismantle", 46200],
    ["Vendor B", "2025", "REGION 1", "RAN & Power Supply", 70500],
    ["Vendor B", "2025", "REGION 1", "TOTAL", 124900],

    ["Vendor B", "2026", "REGION 1", "Site Survey", 8700],
    ["Vendor B", "2026", "REGION 1", "DG Dismantle", 48900],
    ["Vendor B", "2026", "REGION 1", "RAN & Power Supply", 71900],
    ["Vendor B", "2026", "REGION 1", "TOTAL", 129500],

    ["Vendor B", "TOTAL", "REGION 1", "", 254400],

    ["Vendor C", "2025", "REGION 1", "Site Survey", 8900],
    ["Vendor C", "2025", "REGION 1", "DG Dismantle", 43500],
    ["Vendor C", "2025", "REGION 1", "RAN & Power Supply", 73800],
    ["Vendor C", "2025", "REGION 1", "TOTAL", 126200],

    ["Vendor C", "2026", "REGION 1", "Site Survey", 9700],
    ["Vendor C", "2026", "REGION 1", "DG Dismantle", 51300],
    ["Vendor C", "2026", "REGION 1", "RAN & Power Supply", 69500],
    ["Vendor C", "2026", "REGION 1", "TOTAL", 130500],

    ["Vendor C", "TOTAL", "REGION 1", "", 256700],

    ["Vendor A", "2025", "REGION 2", "Site Survey", 9000],
    ["Vendor A", "2025", "REGION 2", "DG Dismantle", 47800],
    ["Vendor A", "2025", "REGION 2", "RAN & Power Supply", 65500],
    ["Vendor A", "2025", "REGION 2", "TOTAL", 122300],

    ["Vendor A", "2026", "REGION 2", "Site Survey", 8800],
    ["Vendor A", "2026", "REGION 2", "DG Dismantle", 44000],
    ["Vendor A", "2026", "REGION 2", "RAN & Power Supply", 72400],
    ["Vendor A", "2026", "REGION 2", "TOTAL", 125200],

    ["Vendor A", "TOTAL", "REGION 2", "", 247500],

    ["Vendor B", "2025", "REGION 2", "Site Survey", 8700],
    ["Vendor B", "2025", "REGION 2", "DG Dismantle", 49000],
    ["Vendor B", "2025", "REGION 2", "RAN & Power Supply", 67900],
    ["Vendor B", "2025", "REGION 2", "TOTAL", 125600],

    ["Vendor B", "2026", "REGION 2", "Site Survey", 8300],
    ["Vendor B", "2026", "REGION 2", "DG Dismantle", 46800],
    ["Vendor B", "2026", "REGION 2", "RAN & Power Supply", 69300],
    ["Vendor B", "2026", "REGION 2", "TOTAL", 124400],

    ["Vendor B", "TOTAL", "REGION 2", "", 250000],

    ["Vendor C", "2025", "REGION 2", "Site Survey", 9200],
    ["Vendor C", "2025", "REGION 2", "DG Dismantle", 46000],
    ["Vendor C", "2025", "REGION 2", "RAN & Power Supply", 71500],
    ["Vendor C", "2025", "REGION 2", "TOTAL", 126700],

    ["Vendor C", "2026", "REGION 2", "Site Survey", 9300],
    ["Vendor C", "2026", "REGION 2", "DG Dismantle", 49100],
    ["Vendor C", "2026", "REGION 2", "RAN & Power Supply", 66800],
    ["Vendor C", "2026", "REGION 2", "TOTAL", 125200],

    ["Vendor C", "TOTAL", "REGION 2", "", 251900],
]
df_summary = pd.DataFrame(data, columns=columns)

num_cols = ["REGION 1", "REGION 2", "TOTAL"]
df_summary_styled = (
    df_summary.style
    .format({col: format_rupiah for col in num_cols})
    .apply(highlight_total_per_year, axis=1)
    .apply(highlight_vendor_total, axis=1)
)

st.dataframe(df_summary_styled, hide_index=True)

st.write("")
st.markdown("**:yellow-badge[3. TCO SUMMARY]**")
st.markdown(
    """
        <div style="text-align: justify; font-size: 15px; margin-bottom: 10px; margin-top:-10px;">
            The system will automatically generate a TCO Summary that includes 
            the TOTAL calculations. Because this case involves a multi-dimensional column structure, 
            the TCO Summary is split into three tabs, as follows.
        </div>
    """,
    unsafe_allow_html=True
)

tab1, tab2, tab3 = st.tabs(["YEAR", "REGION", "SCOPE"])

with tab1:
    # DataFrame
    columns = ["YEAR", "VENDOR A", "VENDOR B", "VENDOR C"]
    data = [
        ["2025", 243800, 250500, 252900],
        ["2026", 255800, 253900, 255700],
        ["TOTAL", 499600, 504400, 508600]
    ]
    df_tco_year = pd.DataFrame(data, columns=columns)

    num_cols = ["VENDOR A", "VENDOR B", "VENDOR C"]
    df_tco_year_styled = (
        df_tco_year.style
        .format({col: format_rupiah for col in num_cols})
        .apply(highlight_total, axis=1)
    )
    st.dataframe(df_tco_year_styled, hide_index=True)

with tab2:
    # DataFrame
    columns = ["REGION", "VENDOR A", "VENDOR B", "VENDOR C"]
    data = [
        ["REGION 1", 252100, 254400, 256700],
        ["REGION 2", 247500, 250000, 251900],
        ["TOTAL", 499600, 504400, 508600]
    ]
    df_tco_region = pd.DataFrame(data, columns=columns)

    num_cols = ["VENDOR A", "VENDOR B", "VENDOR C"]
    df_tco_region_styled = (
        df_tco_region.style
        .format({col: format_rupiah for col in num_cols})
        .apply(highlight_total, axis=1)
    )
    st.dataframe(df_tco_region_styled, hide_index=True)

with tab3:
    # DataFrame
    columns = ["SCOPE", "VENDOR A", "VENDOR B", "VENDOR C"]
    data = [
        ["DG Dismantle", 183300, 190900, 189900],
        ["RAN & Power Supply", 280800, 279600, 281600],
        ["Site Survey", 35500, 33900, 37100],
        ["TOTAL", 499600, 504400, 508600]
    ]
    df_tco_scope = pd.DataFrame(data, columns=columns)

    num_cols = ["VENDOR A", "VENDOR B", "VENDOR C"]
    df_tco_scope_styled = (
        df_tco_scope.style
        .format({col: format_rupiah for col in num_cols})
        .apply(highlight_total, axis=1)
    )
    st.dataframe(df_tco_scope_styled, hide_index=True)

st.write("")
st.markdown("**:green-badge[4. BID & PRICE ANALYSIS]**")
st.markdown(
    """
        <div style="text-align: justify; font-size: 15px; margin-bottom: 10px; margin-top:-10px;">
            This menu also displays an analysis table that provides a comprehensive overview of the pricing structure 
            submitted by each vendor, as follows.
        </div>
    """,
    unsafe_allow_html=True
)

st.markdown(
    """
    <div style="text-align:left; margin-bottom: 8px">
        <span style="background:#C6EFCE; padding:2px 8px; border-radius:6px; font-weight:600; font-size: 0.75rem; color: black">1st Lowest</span>
        &nbsp;
        <span style="background:#FFEB9C; padding:2px 8px; border-radius:6px; font-weight:600; font-size: 0.75rem; color: black">2nd Lowest</span>
    </div>
    """,
    unsafe_allow_html=True
)

# DataFrame
columns = ["YEAR", "REGION", "SCOPE", "VENDOR A", "VENDOR B", "VENDOR C", "1st Lowest", "1st Vendor", "2nd Lowest", "2nd Vendor", "Gap 1 to 2 (%)", "Median Price", "VENDOR A to Median (%)", "VENDOR B to Median (%)", "VENDOR C to Median (%)"]
data = [
    ["2025","REGION 1","DG Dismantle",45000,46200,43500,43500,"VENDOR C",45000,"VENDOR A","3.5%",45000,"+0.0%","+2.7%","-3.3%"],
    ["2025","REGION 1","RAN & Power Supply",68000,70500,73800,68000,"VENDOR A",70500,"VENDOR B","3.7%",70500,"-3.5%","+0.0%","+4.7%"],
    ["2025","REGION 1","Site Survey",8500,8200,8900,8200,"VENDOR B",8500,"VENDOR A","3.7%",8500,"+0.0%","-3.5%","+4.7%"],
    ["2025","REGION 2","DG Dismantle",47800,49000,46000,46000,"VENDOR C",47800,"VENDOR A","3.9%",47800,"+0.0%","+2.5%","-3.8%"],
    ["2025","REGION 2","RAN & Power Supply",65500,67900,71500,65500,"VENDOR A",67900,"VENDOR B","3.7%",67900,"-3.5%","+0.0%","+5.3%"],
    ["2025","REGION 2","Site Survey",9000,8700,9200,8700,"VENDOR B",9000,"VENDOR A","3.5%",9000,"+0.0%","-3.3%","+2.2%"],
    ["2026","REGION 1","DG Dismantle",46500,48900,51300,46500,"VENDOR A",48900,"VENDOR B","5.2%",48900,"-4.9%","+0.0%","+4.9%"],
    ["2026","REGION 1","RAN & Power Supply",74900,71900,69500,69500,"VENDOR C",71900,"VENDOR B","3.5%",71900,"+4.2","+0.0%","-3.3%"],
    ["2026","REGION 1","Site Survey",9200,8700,9700,8700,"VENDOR B",9200,"VENDOR A","5.8%",9200,"+0.0%","-5.4%","+5.4%"],
    ["2026","REGION 2","DG Dismantle",44000,46800,49100,44000,"VENDOR A",46800,"VENDOR B","6.4%",46800,"-6.0%","+0.0%","+4.9%"],
    ["2026","REGION 2","RAN & Power Supply",72400,69300,66800,66800,"VENDOR C",69300,"VENDOR B","3.7%",69300,"+4.5%","+0.0%","-3.6%"],
    ["2026","REGION 2","Site Survey",8800,8300,9300,8300,"VENDOR B",8800,"VENDOR A","6.0%",8800,"+0.0%","-5.7%","+5.7%"],
]
df_analysis = pd.DataFrame(data, columns=columns)

num_cols = ["VENDOR A", "VENDOR B", "VENDOR C", "1st Lowest", "2nd Lowest", "Median Price"]
df_analysis_styled = (
    df_analysis.style
    .format({col: format_rupiah for col in num_cols})
    .apply(lambda row: highlight_1st_2nd_vendor(row, df_analysis.columns), axis=1)
)

st.dataframe(df_analysis_styled, hide_index=True)

st.write("")
st.markdown("**:blue-badge[5. VISUALIZATION]**")
st.markdown(
    """
        <div style="text-align: justify; font-size: 15px; margin-bottom: 10px; margin-top:-10px;">
            This menu displays visualizations focusing on two key aspects: <span style="background: #FF5E5E; 
            padding:1px 4px; border-radius:6px; font-weight:600; font-size: 13px; color: black">Win Rate Trend</span> 
            and <span style="background: #FF00AA; padding:2px 4px; border-radius:6px; font-weight:600; 
            font-size: 13px; color: black">Average Gap Trend</span>, each presented in its own tab.
        </div>
    """,
    unsafe_allow_html=True
)

tab1, tab2 = st.tabs(["Win Rate Trend", "Average Gap Trend"])

with tab1:
    st.image("assets/1.png")
    with st.expander("See explanation"):
        st.caption('''
            The visualization above compares the win rate of each vendor
            based on how often they achieved 1st or 2nd place in all
            tender evaluations.  
                    
            **💡 How to interpret the chart**  
                    
            - High 1st Win Rate (%)  
                Vendor is highly competitive and often offers the best commercial terms.  
            - High 2nd Win Rate (%)  
                Vendor consistently performs well, often just slightly less competitive than the winner.  
            - Large Gap Between 1st & 2nd Win Rate  
                Shows clear market dominance by certain vendors.
        ''')

with tab2:
    st.image("assets/2.png")
    with st.expander("See explanation"):
        st.caption('''
            The chart above shows the average price difference between 
            the lowest and second-lowest bids for each vendor when they 
            rank 1st, indicating their pricing dominance or competitiveness.
                    
            **💡 How to interpret the chart**  
                    
            - High Gap  
                High gap indicates strong vendor dominance (much lower prices).  
            - Low Gap  
                Low gap indicates intense competition with similar pricing among vendors.  
            
            The dashed line represents the average gap across all vendors, serving as a benchmark (4.4%).
        ''')
    
st.write("")
st.markdown("**:violet-badge[6. SUPER BUTTON]**")
st.markdown(
    """
        <div style="text-align: justify; font-size: 15px; margin-bottom: 10px; margin-top:-10px;">
            Lastly, there is a <span style="background:#FFCB09; padding:2px 4px; border-radius:6px; font-weight:600; 
            font-size: 0.75rem; color: black">Super Button</span> feature where all dataframes generated by the system 
            can be downloaded as a single file with multiple sheets. You can also customize the order of the sheets.
            The interface looks more or less like this.
        </div>
    """,
    unsafe_allow_html=True
)

dataframes = {
    "Merge Data": df_merge,
    "Cost Summary": df_summary,
    "TCO Summary (Year)": df_tco_year,
    "TCO Summary (Region)": df_tco_region,
    "TCO Summary (Scope)": df_tco_scope,
    "Bid & Price Analysis": df_analysis,
}

# Tampilkan multiselect
selected_sheets = st.multiselect(
    "Select sheets to download in a single Excel file:",
    options=list(dataframes.keys()),
    default=list(dataframes.keys())  # default semua dipilih
)

# Fungsi "Super Button" & Formatting
def generate_multi_sheet_excel(selected_sheets, df_dict):
    """
    Gabungkan sheet dengan highlight khusus per sheet:
    - Merge Data / Cost Summary: highlight TOTAL per year & TOTAL vendor
    - Bid & Price Analysis: highlight 1st & 2nd vendor + TOTAL
    - TCO sheets: highlight TOTAL
    Semua akses baris pakai index, jadi aman untuk kolom dengan spasi/simbol.
    """
    output = BytesIO()

    with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
        for sheet in selected_sheets:
            df = df_dict[sheet].copy()
            workbook  = writer.book
            worksheet = workbook.add_worksheet(sheet)
            writer.sheets[sheet] = worksheet

            # --- Format umum ---
            fmt_rupiah = workbook.add_format({'num_format': '#,##0'})
            fmt_pct    = workbook.add_format({'num_format': '#,##0.0"%"'})
            fmt_total  = workbook.add_format({'bold': True, 'bg_color': '#D9EAD3', 'font_color': '#1A5E20', 'num_format': '#,##0'})
            fmt_first  = workbook.add_format({'bg_color': '#C6EFCE', "num_format": "#,##0"})
            fmt_second = workbook.add_format({'bg_color': '#FFEB9C', "num_format": "#,##0"})
            fmt_total_year  = workbook.add_format({'bold': True, 'bg_color': '#FFEB9C', 'font_color': '#1A1A1A', 'num_format': '#,##0'})
            fmt_total_vendor  = workbook.add_format({'bold': True, 'bg_color': '#C6EFCE', 'font_color': '#1A5E20', 'num_format': '#,##0'})

            # --- Tulis header ---
            for col_idx, col_name in enumerate(df.columns):
                worksheet.write(0, col_idx, col_name)

            numeric_cols = df.select_dtypes(include=["number"]).columns.tolist()
            vendor_cols  = [c for c in numeric_cols] if sheet == "Bid & Price Analysis" else []

            # Atur lebar kolom
            for col_idx, col_name in enumerate(df.columns):
                if col_name in numeric_cols:
                    worksheet.set_column(col_idx, col_idx, 15, fmt_rupiah)
                if "%" in col_name:
                    worksheet.set_column(col_idx, col_idx, 15, fmt_pct)

            # --- Tulis baris dengan highlight ---
            for row_idx, row in enumerate(df.itertuples(index=False), start=1):
                fmt_row = [None]*len(df.columns)

                # --- Merge Data highlight ---
                if sheet == "Merge Data":
                    year_val  = str(row[1]).upper()   # kolom Year
                    scope_val = str(row[2]).upper()   # kolom Scope
                    if scope_val == "TOTAL" and year_val != "TOTAL":
                        fmt_row = [fmt_total_year]*len(df.columns)
                    elif year_val == "TOTAL":
                        fmt_row = [fmt_total_vendor]*len(df.columns)

                # --- Cost Summary highlight ---
                elif sheet == "Cost Summary":
                    year_val  = str(row[1]).upper()
                    scope_val = str(row[3]).upper()
                    if scope_val == "TOTAL" and year_val != "TOTAL":
                        fmt_row = [fmt_total_year]*len(df.columns)
                    elif year_val == "TOTAL":
                        fmt_row = [fmt_total_vendor]*len(df.columns)

                # --- Bid & Price Analysis highlight ---
                elif sheet == "Bid & Price Analysis":
                    first_vendor_name  = row[df.columns.get_loc("1st Vendor")]
                    second_vendor_name = row[df.columns.get_loc("2nd Vendor")]
                    for col_idx2, col_name2 in enumerate(df.columns):
                        if col_name2 == first_vendor_name:
                            fmt_row[col_idx2] = fmt_first
                        elif col_name2 == second_vendor_name:
                            fmt_row[col_idx2] = fmt_second
                        elif str(row[col_idx2]).upper() == "TOTAL":
                            fmt_row[col_idx2] = fmt_total

                # --- TCO sheets highlight TOTAL ---
                else:
                    if any(str(row[i]).upper() == "TOTAL" for i in range(len(df.columns)) if row[i] is not None):
                        fmt_row = [fmt_total]*len(df.columns)

                # --- Tulis sel ---
                for col_idx2 in range(len(df.columns)):
                    value = row[col_idx2]
                    if pd.isna(value) or (isinstance(value, (int,float)) and np.isinf(value)):
                        value = ""
                    worksheet.write(row_idx, col_idx2, value, fmt_row[col_idx2] if fmt_row[col_idx2] else None)

    output.seek(0)
    return output.getvalue()

# ---- DOWNLOAD BUTTON ----
if selected_sheets:
    excel_bytes = generate_multi_sheet_excel(selected_sheets, dataframes)

    st.download_button(
        label="Download",
        data=excel_bytes,
        file_name="super botton.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        type="primary",
        use_container_width=True,
    )

st.write("")
st.divider()

st.markdown("#### Video Tutorial")
st.markdown(
    """
        <div style="text-align: justify; font-size: 15px; margin-bottom: 10px; margin-top:-10px;">
            I have also included a video tutorial, which you can access through the 
            <span style="background:#FF0000; padding:2px 4px; border-radius:6px; font-weight:600; 
            font-size: 0.75rem; color: black">YouTube</span> link below.
        </div>
    """,
    unsafe_allow_html=True
)

st.video("https://youtu.be/V2G8ESoDXm8?si=NqNCQxpGWmkvgDBV")