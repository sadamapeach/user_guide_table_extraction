import streamlit as st
import pandas as pd
import numpy as np
import time
import re
from io import BytesIO

def round_half_up(series):
    return np.floor(series * 100 + 0.5) / 100

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

st.subheader("🧑‍🏫 User Guide: Table Extraction")
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
            <span style="color: #0073FF; font-weight: 800;">
            Table Extraction</span>
            is used to extract tables from multi-sheet files where each sheet may contain multiple tables,
            arranged either horizontally or vertically.
        </div>
    </div>
    """,
    unsafe_allow_html=True
)

st.markdown("#### Input Structure")

st.markdown(
    """
        <div style="text-align: justify; font-size: 15px; margin-bottom: 20px">
            The input file required for this menu should be a 
            <span style="color: #FF69B4; font-weight: 500;">single file containing multiple sheets</span>, in eather 
            <span style="background:#C6EFCE; padding:1px 4px; border-radius:6px; font-weight:600; font-size: 0.75rem; color: black">.xlsx</span> or 
            <span style="background:#FFEB9C; padding:2px 4px; border-radius:6px; font-weight:600; font-size: 0.75rem; color: black">.xls</span> format. 
            Each sheet represents a vendor name, with the table structure in each sheet as follows:
        </div>
    """,
    unsafe_allow_html=True
)

# Dataframe
columns = ["A", "B", "C", "D", "E", "F", "G", "H", "I"]
df = pd.DataFrame([[""] * len(columns) for _ in range(8)], columns=columns)

# Table1
df.loc[1, ["B"]] = ["TOTAL TCO-3Y"]
df.loc[2, ["B"]] = ["28.000"]

# Table2
df.loc[4, ["B", "C", "D"]] = ["SCOPE TOTAL", "REGION 1", "REGION 2"]
df.loc[5, ["B", "C", "D"]] = ["MBTS", "4.000", "3.000"]
df.loc[6, ["B", "C", "D"]] = ["Antenna", "6.000", "5.000"]

# Table3
df.loc[4, ["F", "G", "H"]] = ["SCOPE UNIT", "REGION 1", "REGION 2"]
df.loc[5, ["F", "G", "H"]] = ["MBTS", "8.000", "7.000"]
df.loc[6, ["F", "G", "H"]] = ["Antenna", "2.500", "2.300"]

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

st.markdown(
    """
        <div style="text-align: justify; font-size: 15px; margin-bottom: 20px">
            The system accommodates a 
            <span style="font-weight: bold;">dynamic concept</span>, where a single sheet may contain multiple tables. 
            Each table can consist of multiple rows and columns. The table extraction process in this system is capable of handling
            <span style="color: #FF69B4; font-weight: 500;">vertical, horizontal, or mixed table formats,</span> as illustrated in the example above.
        </div>
    """,
    unsafe_allow_html=True
)

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

st.markdown("**:red-badge[1. TABLE SEPARATOR]**")
st.markdown(
    """
        <div style="text-align: justify; font-size: 15px; margin-bottom: 10px; margin-top:-10px;">
            The separator used to distinguish each table is 
            <span style="color: #FF69B4; font-weight: 500;">one or more empty rows</span> and or 
            <span style="color: #FF69B4; font-weight: 500;">empty columns</span>, as shown in the table example in the 
            '<span style="font-weight: bold;">Input Structure</span>' section.
            Please refer to the following example.
        </div>
    """,
    unsafe_allow_html=True
)

# Dataframe
columns = ["A", "B", "C", "D", "E", "F", "G", "H"]
df = pd.DataFrame([[""] * len(columns) for _ in range(8)], columns=columns)

# Table1
df.loc[1, ["B", "C"]] = ["Desc", "Price"]
df.loc[2, ["B", "C"]] = ["Instalasi OTB", "1.300"]
df.loc[3, ["B", "C"]] = ["Instalasi Modem", "1.200"]

# Table2
df.loc[4, ["B", "C"]] = ["Scope", "Price"]
df.loc[5, ["B", "C"]] = ["Ethernet Cable", "4.000"]
df.loc[6, ["B", "C"]] = ["Optical Cable", "3.500"]

# Table3
df.loc[1, ["E", "F", "G"]] = ["Scope Unit", "Region 1", "Region 2"]
df.loc[2, ["E", "F", "G"]] = ["MBTS", "8.000", "7.000"]
df.loc[3, ["E", "F", "G"]] = ["Antenna", "2.500", "2.300"]

st.dataframe(df, hide_index=True)

st.markdown(
    """
        <div style="text-align: justify; font-size: 15px; margin-bottom: 25px; margin-top: -5px;">
            The table above consists of three tables: 
            <span style="background:#FFCB09; padding:2px 4px; border-radius:6px; font-weight:600; font-size: 0.75rem; color: black">Desc</span> 
            <span style="background:#FFCB09; padding:2px 4px; border-radius:6px; font-weight:600; font-size: 0.75rem; color: black">Scope</span> and
            <span style="background:#FFCB09; padding:2px 4px; border-radius:6px; font-weight:600; font-size: 0.75rem; color: black">Scope Unit</span>. 
            However, the 
            <span style="color: #FF69B4; font-weight: bold;">Desc</span> and 
            <span style="color: #FF69B4; font-weight: bold;">Scope</span> tables do not have an empty-row separator, so they will be treated as a single combined table. Meanwhile, the 
            <span style="color: #FF69B4; font-weight: bold;">Scope Unit</span> table has an empty-column separator, so it will be recognized as a separate table.
        </div>
    """,
    unsafe_allow_html=True
)

st.markdown("**:orange-badge[2. DATA TYPES]**")
st.markdown(
    """
        <div style="text-align: justify; font-size: 15px; margin-bottom: 10px; margin-top:-10px;">
            Data types affect both the display format and the download output. Please refer to the example below.
        </div>
    """,
    unsafe_allow_html=True
)

# Dataframe
columns = ["A", "B", "C", "D", "E", "F", "G"]
df = pd.DataFrame([[""] * len(columns) for _ in range(5)], columns=columns)

# Table1
df.loc[1, ["B", "C"]] = ["Desc", "Price"]
df.loc[2, ["B", "C"]] = ["Instalasi OTB", "Discount 10%"]
df.loc[3, ["B", "C"]] = ["Instalasi Modem", "1.200"]

# Table2
df.loc[1, ["E", "F"]] = ["Scope", "Price"]
df.loc[2, ["E", "F"]] = ["Ethernet Cable", "4.000"]
df.loc[3, ["E", "F"]] = ["Optical Cable", "3.500"]

st.dataframe(df, hide_index=True)

st.markdown(
    """
        <div style="text-align: justify; font-size: 15px; margin-bottom: 25px; margin-top:-10px;">
            There are two tables: 
            <span style="background:#FFCB09; padding:2px 4px; border-radius:6px; font-weight:600; font-size: 0.75rem; color: black">Desc</span> & 
            <span style="background:#FFCB09; padding:2px 4px; border-radius:6px; font-weight:600; font-size: 0.75rem; color: black">Scope</span> and both contain 
            <span style="background: #FF5E5E; padding:2px 4px; border-radius:6px; font-weight:600; font-size: 13px; color: black">PRICE COLUMN</span>.
            In the 
            <span style="color: #FF69B4; font-weight: 500;">Desc</span> table, the price column contains mixed data types (non-numeric and numeric), so currency formatting will not be applied.
            In contrast, the 
            <span style="color: #FF69B4; font-weight: 500;">Scope</span> table contains only numeric values. Therefore, currency formatting will be applied correctly. 
        </div>
    """,
    unsafe_allow_html=True
)

st.markdown("**:green-badge[3. FLOATING TABLE]**")
st.markdown(
    """
        <div style="text-align: justify; font-size: 15px; margin-bottom: 10px; margin-top:-10px;">
            Floating tables are allowed, meaning tables 
            <span style="color: #FF69B4; font-weight: 700;">do not need to start from cell A1</span>. 
            However, ensure that the cells above and to the left of the table are empty, as shown in the table example in the 
            '<span style="font-weight: bold;">Input Structure</span>' section.
            To provide additional explanations or notes on the sheet, you can include them using an image or a text box.
        </div>
    """,
    unsafe_allow_html=True
)

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
    file_name="Dummy Dataset - Table Extraction.xlsx",
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

st.markdown("**:blue-badge[1. TABLE EXTRACTION]**")
st.markdown(
    """
    <div style="text-align: justify; font-size: 15px; margin-bottom: 20px">
        The system will extract multiple tables into separate single tables. Each sheet will be
        processed in a loop and displayed as an individual tab, as shown in the example below.
    </div>
    """,
    unsafe_allow_html=True
)

tab1, tab2 = st.tabs(["VENDOR A", "VENDOR B"])

with tab1:
    # Tabel 1
    st.markdown(
        """
        <div style='display: flex; justify-content: space-between; align-items: center; margin-bottom: 8px;'>
            <span style='font-size:14px;'>✨ VENDOR A - Table 1</span>
            <span style='font-size:12px; color:#808080;'>
                Total rows: <b>1</b>
            </span>
        </div>
        """,
        unsafe_allow_html=True
    )

    # Dataframe
    columns = ["TOTAL TCO-3Y"]
    data = [28000]
    df_va_t1 = pd.DataFrame(data, columns=columns)

    num_cols = ["TOTAL TCO-3Y"]
    df_va_t1_styled = (
        df_va_t1.style
        .format({col: format_rupiah for col in num_cols})
    )
    st.dataframe(df_va_t1_styled, hide_index=True)

    # Tabel 2
    st.markdown(
        """
        <div style='display: flex; justify-content: space-between; align-items: center; margin-bottom: 8px;'>
            <span style='font-size:14px;'>✨ VENDOR A - Table 2</span>
            <span style='font-size:12px; color:#808080;'>
                Total rows: <b>2</b>
            </span>
        </div>
        """,
        unsafe_allow_html=True
    )

    # Dataframe
    columns = ["SCOPE TOTAL PRICE", "REGION 1", "REGION 2"]
    data = [
        ["MBTS", 4000, 3000],
        ["Antenna", 6000, 5000],
    ]
    df_va_t2 = pd.DataFrame(data, columns=columns)

    num_cols = ["REGION 1", "REGION 2"]
    df_va_t2_styled = (
        df_va_t2.style
        .format({col: format_rupiah for col in num_cols})
    )
    st.dataframe(df_va_t2_styled, hide_index=True)

    # Tabel 3
    st.markdown(
        """
        <div style='display: flex; justify-content: space-between; align-items: center; margin-bottom: 8px;'>
            <span style='font-size:14px;'>✨ VENDOR A - Table 3</span>
            <span style='font-size:12px; color:#808080;'>
                Total rows: <b>2</b>
            </span>
        </div>
        """,
        unsafe_allow_html=True
    )

    # Dataframe
    columns = ["SCOPE UNIT PRICE", "REGION 1", "REGION 2"]
    data = [
        ["MBTS", 8000, 7000],
        ["Antenna", 2500, 2300],
    ]
    df_va_t3 = pd.DataFrame(data, columns=columns)

    num_cols = ["REGION 1", "REGION 2"]
    df_va_t3_styled = (
        df_va_t3.style
        .format({col: format_rupiah for col in num_cols})
    )
    st.dataframe(df_va_t3_styled, hide_index=True)

with tab2:
    # Tabel 1
    st.markdown(
        """
        <div style='display: flex; justify-content: space-between; align-items: center; margin-bottom: 8px;'>
            <span style='font-size:14px;'>✨ VENDOR B - Table 1</span>
            <span style='font-size:12px; color:#808080;'>
                Total rows: <b>2</b>
            </span>
        </div>
        """,
        unsafe_allow_html=True
    )

    # Dataframe
    columns = ["Description", "Price (IDR)"]
    data = [
        ["Instalasi OTB", 1300],
        ["Instalasi Modem", 1200],
    ]
    df_vb_t1 = pd.DataFrame(data, columns=columns)

    num_cols = ["Price (IDR)"]
    df_vb_t1_styled = (
        df_vb_t1.style
        .format({col: format_rupiah for col in num_cols})
    )
    st.dataframe(df_vb_t1_styled, hide_index=True)

    # Tabel 2
    st.markdown(
        """
        <div style='display: flex; justify-content: space-between; align-items: center; margin-bottom: 8px;'>
            <span style='font-size:14px;'>✨ VENDOR B - Table 2</span>
            <span style='font-size:12px; color:#808080;'>
                Total rows: <b>2</b>
            </span>
        </div>
        """,
        unsafe_allow_html=True
    )

    # Dataframe
    columns = ["Scope", "Total Price (IDR)"]
    data = [
        ["Ethernet Cable", 4000],
        ["Optical Cable", 3500],
    ]
    df_vb_t2 = pd.DataFrame(data, columns=columns)

    num_cols = ["Total Price (IDR)"]
    df_vb_t2_styled = (
        df_vb_t2.style
        .format({col: format_rupiah for col in num_cols})
    )
    st.dataframe(df_vb_t2_styled, hide_index=True)
    
st.write("")
st.markdown("**:violet-badge[2. SUPER BUTTON]**")
st.markdown(
    """
        <div style="text-align: justify; font-size: 15px; margin-bottom: 10px; margin-top:-10px;">
            Lastly, there is a 
            <span style="background:#FFCB09; padding:2px 4px; border-radius:6px; font-weight:600; font-size: 0.75rem; color: black">Super Button</span> 
            feature where all dataframes generated by the system can be downloaded as a single file with multiple sheets. 
            You can also customize the order of the sheets.
            The interface looks more or less like this.
        </div>
    """,
    unsafe_allow_html=True
)

all_sheets_tables = {
    "VENDOR A": [df_va_t1, df_va_t2, df_va_t3],
    "VENDOR B": [df_vb_t1, df_vb_t2],
}

# Tampilkan multiselect
selected_sheets = st.multiselect(
    "Select sheets to download in a single Excel file:",
    options=list(all_sheets_tables.keys()),
    default=list(all_sheets_tables.keys())
)

# Fungsi "Super Button" & Formatting
def to_excel(dfs_dict):
    output = BytesIO()

    with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
        workbook = writer.book
        fmt_rp = workbook.add_format({"num_format": "#,##0"})

        for sheet_name, table_list in dfs_dict.items():
            for idx, df in enumerate(table_list, start=1):
                df = df.copy()

                num_cols = df.select_dtypes(include="number").columns.tolist()
                for col in num_cols:
                    df[col] = round_half_up(df[col])

                tab_name = f"{sheet_name}_Table{idx}"
                df.to_excel(writer, index=False, sheet_name=tab_name)
                worksheet = writer.sheets[tab_name]

                # === AUTOFIT + FORMAT ===
                for i, col in enumerate(df.columns):
                    max_len = max(
                        len(str(col)),
                        df[col].astype(str).map(len).max()
                    ) + 2

                    worksheet.set_column(
                        i, i, max_len, fmt_rp if col in num_cols else None
                    )

    output.seek(0)
    return output.getvalue()

# --- FRAGMENT UNTUK BALLOONS ---
@st.fragment
def release_the_balloons():
    st.balloons()

if selected_sheets:
    dfs_to_export = {
        k: v for k, v in all_sheets_tables.items() if k in selected_sheets
    }

    excel_data = to_excel(dfs_to_export)

    st.download_button(
        label="Download",
        data=excel_data,
        file_name="Super Botton - Table Extraction.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        on_click=release_the_balloons,
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
            <span style="background:#FF0000; padding:2px 4px; border-radius:6px; font-weight:600; font-size: 0.75rem; color: black">YouTube</span> 
            link below.
        </div>
    """,
    unsafe_allow_html=True
)

st.video("https://youtu.be/_kqg84j2t-k?si=jpM6hcuqy5udC1Zc")