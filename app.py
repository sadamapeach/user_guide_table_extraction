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
st.write("On progress..")

# st.markdown(
#     """
#     <div style="
#         display: flex;
#         align-items: center;
#         height: 65px;
#         margin-bottom: 10px;
#     ">
#         <div style="text-align: justify; font-size: 15px;">
#             <span style="color: #0073FF; font-weight: 800;">
#             Table Extraction</span>
#             is used to extract tables from multi-sheet files where each sheet may contain multiple tables,
#             arranged either horizontally or vertically.
#         </div>
#     </div>
#     """,
#     unsafe_allow_html=True
# )

# st.markdown("#### Input Structure")

# st.markdown(
#     """
#         <div style="text-align: justify; font-size: 15px; margin-bottom: 20px">
#             The input file required for this menu should be a 
#             <span style="color: #FF69B4; font-weight: 500;">single file containing multiple sheets</span>, in eather 
#             <span style="background:#C6EFCE; padding:1px 4px; border-radius:6px; font-weight:600; font-size: 0.75rem; color: black">.xlsx</span> or 
#             <span style="background:#FFEB9C; padding:2px 4px; border-radius:6px; font-weight:600; font-size: 0.75rem; color: black">.xls</span> format. 
#             Each sheet represents a vendor name, with the table structure in each sheet as follows:
#         </div>
#     """,
#     unsafe_allow_html=True
# )

# # Dataframe
# columns = ["", "A", "B", "C", "D", "E", "F", "G", "H", "I"]
# df = pd.DataFrame([[i] + [""] * 9 for i in range(1, 9)], columns=columns)

# # Table1
# df.loc[1, ["B"]] = ["TOTAL TCO-3Y"]
# df.loc[2, ["B"]] = ["28.730.663"]

# # Table2
# df.loc[4, ["B", "C", "D"]] = ["SCOPE TOTAL", "REGION 1", "REGION 2"]
# df.loc[5, ["B", "C", "D"]] = ["MBTS", "409.430", "304.257"]
# df.loc[6, ["B", "C", "D"]] = ["Antenna", "690.316", "550.101"]

# # Table3
# df.loc[4, ["F", "G", "H"]] = ["SCOPE UNIT", "REGION 1", "REGION 2"]
# df.loc[5, ["F", "G", "H"]] = ["MBTS", "81.886", "76.064"]
# df.loc[6, ["F", "G", "H"]] = ["Antenna", "25.567", "23.917"]

# st.dataframe(df, hide_index=True)

# # Buat DataFrame 1 row
# st.markdown("""
# <table style="width: 100%; border-collapse: collapse; table-layout: fixed; font-size: 15px;">
#     <tr>
#         <td style="border: 1px solid gray; width: 15%;">Vendor A</td>
#         <td style="border: 1px solid gray; width: 15%;">Vendor B</td>
#         <td style="border: 1px solid gray; width: 15%;">Vendor C</td>
#         <td style="border: 1px solid gray; font-style: italic; color: #26BDAD">multiple sheets</td>
#     </tr>
# </table>
# """, unsafe_allow_html=True)

# st.markdown("###### Description:")
# st.markdown(
#     """
#     <div style="font-size:15px;">
#         <ul>
#             <li>
#                 <span style="display:inline-block; width:100px;">Scope & Desc</span>: non-numeric columns
#             </li>
#             <li>
#                 <span style="display:inline-block; width:100px;">Y0 to Y5</span>: numeric columns
#             </li>
#             <li>
#                 <span style="display:inline-block; width:100px;">TOTAL 5Y</span>: optional column
#             </li>
#         </ul>
#     </div>
#     """,
#     unsafe_allow_html=True
# )

# st.markdown(
#     """
#         <div style="text-align: justify; font-size: 15px; margin-bottom: 20px">
#             The system accommodates a 
#             <span style="font-weight: bold;">dynamic table</span>, allowing users to enter any number of non-numeric and numeric columns. The 
#             <span style="background:#FFCB09; padding:2px 4px; border-radius:6px; font-weight:600; font-size: 0.75rem; color: black">TOTAL 5Y</span> 
#             column is optional and can be included or omitted. Users have the freedom to name the columns as they wish. The system logic relies on 
#             <span style="font-weight: bold;">column indices</span>, not specific column names.
#         </div>
#     """,
#     unsafe_allow_html=True
# )

# st.markdown("**:violet-badge[Ensure that each sheet has the same table structure and column names!]**")

# st.divider()
# st.markdown("#### Constraint")

# st.markdown(
#     """
#         <div style="text-align: justify; font-size: 15px; margin-bottom: 20px; margin-top: -10px">
#             To ensure this menu works correctly, users need to follow certain rules regarding
#             the dataset structure.
#         </div>
#     """,
#     unsafe_allow_html=True
# )

# st.markdown("**:red-badge[1. COLUMN ORDER]**")
# st.markdown(
#     """
#         <div style="text-align: justify; font-size: 15px; margin-bottom: 10px; margin-top: -10px">
#             When creating tables, it is important to follow the specified column structure. Columns 
#             <span style="font-weight: bold;">must</span> be arranged in the following order:
#         </div>
#     """,
#     unsafe_allow_html=True
# )

# st.markdown(
#     """
#         <div style="text-align: center; font-size: 15px; margin-bottom: 10px; font-weight: bold">
#             Non-Numeric Columns → Numeric Columns → TOTAL Column (optional)
#         </div>
#     """,
#     unsafe_allow_html=True
# )

# st.markdown(
#     """
#         <div style="text-align: justify; font-size: 15px; margin-bottom: 25px">
#             this order is <span style="color: #FF69B4; font-weight: 700;">strict</span> and 
#             <span style="color: #FF69B4; font-weight: 700;">cannot be altered</span>!
#         </div>
#     """,
#     unsafe_allow_html=True
# )

# st.markdown("**:orange-badge[2. NUMBER COLUMN]**")
# st.markdown(
#     """
#         <div style="text-align: justify; font-size: 15px; margin-bottom: 10px; margin-top:-10px;">
#             Please refer the table below:
#         </div>
#     """,
#     unsafe_allow_html=True
# )

# # DataFrame
# columns = ["No", "Scope", "Desc", "Y0", "Y1", "Y2", "Y3", "Y4", "Y5", "TOTAL 5Y"]
# data = [
#     [1] + [""] * (len(columns) - 1),
#     [2] + [""] * (len(columns) - 1),
#     [3] + [""] * (len(columns) - 1)
# ]
# df = pd.DataFrame(data, columns=columns)

# st.dataframe(df, hide_index=True)

# st.markdown(
#     """
#         <div style="text-align: justify; font-size: 15px; margin-bottom: 25px; margin-top: -5px;">
#             The table above is an 
#             <span style="color: #FF69B4; font-weight: 700;">incorrect example</span> and is 
#             <span style="color: #FF69B4; font-weight: 700;">not allowed</span> because it contains a 
#             <span style="font-weight: bold;">"No"</span> column. 
#             The "No" column is prohibited in this menu, as it will be treated as a numeric column by the system, 
#             which violates the constraint described in point 1.
#         </div>
#     """,
#     unsafe_allow_html=True
# )

# st.markdown("**:green-badge[3. FLOATING TABLE]**")
# st.markdown(
#     """
#         <div style="text-align: justify; font-size: 15px; margin-bottom: 10px; margin-top:-10px;">
#             Floating tables are allowed, meaning tables 
#             <span style="color: #FF69B4; font-weight: 700;">do not need to start from cell A1</span>. 
#             However, ensure that the cells above and to the left of the table are empty, as shown in the example below:
#         </div>
#     """,
#     unsafe_allow_html=True
# )

# # DataFrame
# columns = ["", "A", "B", "C", "D", "E", "F", "G", "H"]
# df = pd.DataFrame([[i] + [""] * 8 for i in range(1, 6)], columns=columns)

# # Header row
# df.loc[1, ["B", "C", "D", "E", "F", "G"]] = ["Scope", "Y0", "Y1", "Y2", "Y3", "TOTAL 3Y TCO"]

# # Data rows
# df.loc[2, ["B", "C", "D", "E", "F", "G"]] = ["Software", "1.000", "2.000", "3.000", "4.000", "10.000"]
# df.loc[3, ["B", "C", "D", "E", "F", "G"]] = ["Hardware", "1.500", "2.500", "3.500", "4.500", "12.000"]

# st.dataframe(df, hide_index=True)

# st.markdown(
#     """
#         <div style="text-align: justify; font-size: 15px; margin-bottom: 25px; margin-top:-10px;">
#             To provide additional explanations or notes on the sheet, you can include them using an image or a text box.
#         </div>
#     """,
#     unsafe_allow_html=True
# )

# st.markdown("**:blue-badge[4. TOTAL ROW]**")
# st.markdown(
#     """
#         <div style="text-align: justify; font-size: 15px; margin-bottom: 10px; margin-top:-10px;">
#             You are not allowed to add a 
#             <span style="font-weight: 700;">TOTAL</span> row at the bottom of the table! 
#             Please refer to the example table below:
#         </div>
#     """,
#     unsafe_allow_html=True
# )

# # DataFrame
# columns = ["Scope", "Y0", "Y1", "Y2", "Y3", "TOTAL 3Y TCO"]
# data = [
#     ["Software", "1.000", "2.000", "3.000", "4.000", "10.000"],
#     ["Hardware", "1.500", "2.500", "3.500", "4.500", "12.000"],
#     ["TOTAL", "2.500", "4.500", "6.500", "8.500", "22.000"],
# ]
# df = pd.DataFrame(data, columns=columns)

# def red_highlight(row):
#     if row["Scope"] == "TOTAL":
#         return ["color: #FF4D4D;" for _ in row]
#     return [""] * len(row)

# df_styled = df.style.apply(red_highlight, axis=1)

# st.dataframe(df_styled, hide_index=True)

# st.markdown(
#     """
#         <div style="text-align: justify; font-size: 15px; margin-bottom: 20px; margin-top: -5px;">
#             The table above is an 
#             <span style="color: #FF69B4; font-weight: 700;">incorrect example</span> and is 
#             <span style="color: #FF69B4; font-weight: 700;">not permitted</span>! 
#             The total row is generated automatically during
#             <span style="font-weight: 700;">MERGE DATA</span> — 
#             do not add one manually, or it will be treated as part of the scope and included in calculations.
#         </div>
#     """,
#     unsafe_allow_html=True
# )

# st.divider()

# st.markdown("#### What is Displayed?")

# # Path file Excel yang sudah ada
# file_path = "dummy dataset.xlsx"

# # Buka file sebagai binary
# with open(file_path, "rb") as f:
#     file_data = f.read()

# # Markdown teks
# st.markdown(
#     """
#     <div style="text-align: justify; font-size: 15px; margin-bottom: 5px; margin-top: -10px">
#         You can try this menu by downloading the dummy dataset using the button below: 
#     </div>
#     """,
#     unsafe_allow_html=True
# )

# @st.fragment
# def release_the_balloons():
#     st.balloons()

# # Download button untuk file Excel
# st.download_button(
#     label="Dummy Dataset",
#     data=file_data,
#     file_name="Dummy Dataset - TCO Comparison by Year.xlsx",
#     mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
#     on_click=release_the_balloons,
#     type="primary",
#     use_container_width=True,
# )

# st.markdown(
#     """
#     <div style="text-align: justify; font-size: 15px; margin-bottom: 20px">
#         Based on this dummy dataset, the menu will produce the following results.
#     </div>
#     """,
#     unsafe_allow_html=True
# )

# st.markdown("**:red-badge[1. MERGE DATA]**")
# st.markdown(
#     """
#         <div style="text-align: justify; font-size: 15px; margin-bottom: 10px; margin-top:-10px;">
#             The system will merge the tables from each sheet into a single table and add a 
#             <span style="background:#FFCB09; padding:2px 4px; border-radius:6px; font-weight:600; font-size: 0.75rem; color: black">TOTAL ROW</span> 
#             for each vendor, as shown below.
#         </div>
#     """,
#     unsafe_allow_html=True
# )

# # DataFrame
# columns = ["VENDOR", "TCO Component", "Y0", "Y1", "Y2", "Y3", "TOTAL"]
# data = [
#     ["Vendor A", "Software", 14200, 15800, 16300, 15500, 61800],
#     ["Vendor A", "Hardware", 20500, 19800, 21000, 22300, 83600],
#     ["Vendor A", "TOTAL", 34700, 35600, 37300, 37800, 145400],
#     ["Vendor B", "Software", 13600, 14200, 14900, 15400, 58100],
#     ["Vendor B", "Hardware", 19300, 20100, 19800, 20600, 79800],
#     ["Vendor B", "TOTAL", 32900, 34300, 34700, 36000, 137900],
#     ["Vendor C", "Software", 14200, 14900, 15600, 16100, 60800],
#     ["Vendor C", "Hardware", 20300, 21000, 21500, 22100, 84900],
#     ["Vendor C", "TOTAL", 34500, 35900, 37100, 38200, 145700],
# ]
# df_merge = pd.DataFrame(data, columns=columns)

# num_cols = ["Y0", "Y1", "Y2", "Y3", "TOTAL"]
# df_merge_styled = (
#     df_merge.style
#     .format({col: format_rupiah for col in num_cols})
#     .apply(highlight_total, axis=1)
# )

# st.dataframe(df_merge_styled, hide_index=True)

# st.write("")
# st.markdown("**:orange-badge[2. TCO SUMMARY]**")
# st.markdown(
#     """
#         <div style="text-align: justify; font-size: 15px; margin-bottom: 10px; margin-top:-10px;">
#             After merging the data, the system will automatically generate a TCO Summary that includes 
#             the TOTAL calculations, as shown below.
#         </div>
#     """,
#     unsafe_allow_html=True
# )

# # DataFrame
# columns = ["TCO Component", "Vendor A", "Vendor B", "Vendor C"]
# data = [
#     ["Software", 61800, 58100, 60800],
#     ["Hardware", 83600, 79800, 84900],
#     ["TOTAL", 145400, 137900, 145700]
# ]
# df_tco = pd.DataFrame(data, columns=columns)

# num_cols = ["Vendor A", "Vendor B", "Vendor C"]
# df_tco_styled = (
#     df_tco.style
#     .format({col: format_rupiah for col in num_cols})
#     .apply(highlight_total, axis=1)
# )
# st.dataframe(df_tco_styled, hide_index=True)

# st.write("")
# st.markdown("**:yellow-badge[3. CURRENCY CONVERTER]**")
# st.markdown(
#     """
#         <div style="text-align: justify; font-size: 15px; margin-bottom: 10px; margin-top:-10px;">
#             This is an additional optional button feature, allowing users to convert the values in the TCO Summary 
#             into other currencies, such as USD, EUR, CNY, IDR, and more. Try performing the conversion using the 
#             button below.
#         </div>
#     """,
#     unsafe_allow_html=True
# )

# # NEW FEATURE: CONVERGE
# # --- Fungsi reset ---
# def reset_fields():
#     for key in ["amount_input", "currency_input", "tco_by_year_amount", "tco_by_year_currency", "converted_tco_by_year", "widget_key"]:
#         if key in st.session_state:
#             del st.session_state[key]
#     # Tambah key unik biar widget benar-benar re-render kosong
#     st.session_state["widget_key"] = str(time.time())

#     # Set flag untuk rerun
#     st.session_state["should_rerun"] = True

# # Jalankan rerun di luar callback (AMAN)
# if st.session_state.get("should_rerun", False):
#     st.session_state["should_rerun"] = False
#     st.rerun()

# # # --- Dapatkan key unik untuk widget ---
# widget_key = st.session_state.get("widget_key", "default")

# # Ambil nilai default (kalau sebelumnya sudah ada)
# default_amount = st.session_state.get("tco_by_year_amount", "")
# default_currency = st.session_state.get("tco_by_year_currency", "")

# # --- Popover ---
# col1, col2, col3 = st.columns([2.3,2,1])
# with col1:
#     with st.popover("Currency Converter"):
#         col1, col2 = st.columns([2, 1])

#         with col1:
#             amount_input = st.text_input(
#                 "Enter amount to convert",
#                 placeholder="e.g. 15000, 0.67",
#                 key=f"amount_input_{widget_key}",  # 🔑 pakai key unik
#                 value=default_amount
#             )

#         with col2:
#             currency_options = ["", "USD", "EUR", "GBP", "SGD", "JPY", "CNY", "INR", "AUD", "CHF", "IDR"]
#             index = currency_options.index(default_currency) if default_currency in currency_options else 0
#             currency_input = st.selectbox(
#                 "Currency",
#                 currency_options,
#                 key=f"currency_input_{widget_key}",  # 🔑 pakai key unik
#                 index=index
#             )

#         st.button("Reset", on_click=reset_fields, type="primary")

# # --- Simpan nilai setelah widget dirender ---
# if amount_input:
#     st.session_state["tco_by_year_amount"] = amount_input
# if currency_input:
#     st.session_state["tco_by_year_currency"] = currency_input

# # --- Ambil kembali nilai ---
# amount = st.session_state.get("tco_by_year_amount", "")
# currency = st.session_state.get("tco_by_year_currency", "")

# # Tampilkan tabel hasil konversi hanya jika input terisi
# if amount and currency:
#     try:
#         # Preprocessing dulu
#         cleaned_amount = re.sub(r"[^\d,\.]", "", amount)
#         if "," in cleaned_amount and "." not in cleaned_amount:
#             cleaned_amount = cleaned_amount.replace(",", ".")

#         # Hapus tanda pemisah ribuan (baik koma maupun titik)
#         cleaned_amount = re.sub(r"(?<=\d)[.,](?=\d{3}(\D|$))", "", cleaned_amount)

#         # Konversi nominal jadi float
#         amount_value = float(cleaned_amount)

#         # Salin merged untuk dikalikan
#         df_converted = df_tco.drop(df_tco.index[-1]).copy(deep=True)

#         # Identifikasi kolom numerik
#         def is_convertible_numeric(series: pd.Series) -> bool:
#             coerced = pd.to_numeric(series, errors="coerce")
#             return coerced.notna().any()

#         # jangan sentuh kolom pertama (biasanya TCO Component)
#         cols_except_first = list(df_converted.columns[1:])

#         # pilih kolom yang "bisa" menjadi numeric dari kolom-kolom tersebut
#         numeric_cols_to_multiply = [
#             c for c in cols_except_first if is_convertible_numeric(df_converted[c])
#         ]

#         # Konversi & kalikan hanya pada kolom terdeteksi
#         if numeric_cols_to_multiply:
#             df_converted.loc[:, numeric_cols_to_multiply] = (
#                 df_converted.loc[:, numeric_cols_to_multiply]
#                 .apply(pd.to_numeric, errors="coerce")
#                 .multiply(amount_value)
#             )

#         # Simpan hasil ke session_state (biar tidak hilang)
#         st.session_state["converted_tco_by_year"] = df_converted

#     except ValueError:
#         st.error("❌ Invalid number format. Please check your input.")

#     # Download button for converter
#     if "converted_tco_by_year" in st.session_state:
#         df_converted = st.session_state["converted_tco_by_year"]
#         currency = st.session_state.get("tco_by_year_currency", "")

#         # Identifikasi numeric & non-numeric columns
#         num_cols = df_converted.select_dtypes(include=["number"]).columns.tolist()
#         non_num_cols = [c for c in df_converted.columns if c not in num_cols]

#         # Buat baris total dinamis
#         total_row = {col: "" for col in df_converted.columns}

#         # Isi label "TOTAL" pada kolom non-numeric pertama
#         if len(non_num_cols) > 0:
#             total_row[non_num_cols[0]] = "TOTAL"

#         # Hitung sum hanya untuk kolom numeric
#         for col in num_cols:
#             total_row[col] = df_converted[col].sum()

#         # Gabungkan
#         df_tco_converted = pd.concat([df_converted, pd.DataFrame([total_row])], ignore_index=True)

#         # Fomat Rupiah & fungsi untuk styling baris TOTAL
#         num_cols_after = df_tco_converted.select_dtypes(include=["number"]).columns

#         converted_styled = (
#             df_tco_converted.style
#             .format({col: format_rupiah for col in num_cols_after})
#             .apply(highlight_total, axis=1)
#         )

#         st.dataframe(converted_styled, hide_index=True)

# st.write("")
# st.markdown("**:green-badge[4. BID & PRICE ANALYSIS]**")
# st.markdown(
#     """
#         <div style="text-align: justify; font-size: 15px; margin-bottom: 10px; margin-top:-10px;">
#             This menu also displays an analysis table that provides a comprehensive overview of the pricing structure 
#             submitted by each vendor, as follows.
#         </div>
#     """,
#     unsafe_allow_html=True
# )

# st.markdown(
#     """
#     <div style="text-align:left; margin-bottom: 8px">
#         <span style="background:#C6EFCE; padding:2px 8px; border-radius:6px; font-weight:600; font-size: 0.75rem; color: black">1st Lowest</span>
#         &nbsp;
#         <span style="background:#FFEB9C; padding:2px 8px; border-radius:6px; font-weight:600; font-size: 0.75rem; color: black">2nd Lowest</span>
#     </div>
#     """,
#     unsafe_allow_html=True
# )

# # DataFrame
# columns = ["TCO Component", "Vendor A", "Vendor B", "Vendor C", "1st Lowest", "1st Vendor", "2nd Lowest", "2nd Vendor", "Gap 1 to 2 (%)", "Median Price", "Vendor A to Median (%)", "Vendor B to Median (%)", "Vendor C to Median (%)"]
# data = [
#     ["Software", 61800, 58100, 60800, 58100, "Vendor B", 60800, "Vendor C", "4.7%", 60800, "+1.6%", "-4.4%", "+0.0%"],
#     ["Hardware", 83600, 79800, 84900, 79800, "Vendor B", 83600, "Vendor A", "4.8%", 83600, "+0.0%", "-4.5%", "+1.6%"],
# ]
# df_analysis = pd.DataFrame(data, columns=columns)

# def highlight_custom(row):
#     # Style
#     yellow = "background-color: #FFEB9C; color: #9C6500;"
#     green = "background-color: #C6EFCE; color: #006100;"
#     default = ""

#     # Mapping posisi sel → style
#     highlight_map = {
#         (1, "Vendor A"): yellow,
#         (0, "Vendor C"): yellow,
#         (0, "Vendor B"): green,
#         (1, "Vendor B"): green,
#     }

#     styled_row = []
#     for col in row.index:
#         cell_coord = (row.name, col)
#         styled_row.append(highlight_map.get(cell_coord, default))

#     return styled_row

# num_cols = ["Vendor A", "Vendor B", "Vendor C", "1st Lowest", "2nd Lowest", "Median Price"]
# df_analysis_styled = (
#     df_analysis.style
#     .format({col: format_rupiah for col in num_cols})
#     .apply(highlight_custom, axis=1)
# )

# st.dataframe(df_analysis_styled, hide_index=True)

# st.write("")
# st.markdown("**:blue-badge[5. RANK VISUALIZATION]**")
# st.markdown(
#     """
#         <div style="text-align: justify; font-size: 15px; margin-bottom: 10px; margin-top:-10px;">
#             The system also displays a rangking visualization consisting of two tabs:
#             <span style="background: #FF5E5E; padding:1px 4px; border-radius:6px; font-weight:600; font-size: 13px; color: black">Original Price</span> and 
#             <span style="background: #FF00AA; padding:2px 4px; border-radius:6px; font-weight:600; font-size: 13px; color: black">Converted Price</span>. 
#             Each tab contains a rank bar chart generated from the
#             <span style="color: #FF69B4; font-weight: 500;">TCO Summary</span> and 
#             <span style="color: #FF69B4; font-weight: 500;">Currency Converter</span> table.
#         </div>
#     """,
#     unsafe_allow_html=True
# )

# tab1, tab2 = st.tabs(["Original Price", "Converted Price"])

# with tab1:
#     # st.image("assets/1.png")
#     st.write("tata")
# with tab2:
#     st.markdown(
#         """
#             <div style="text-align: justify; font-size: 15px; margin-bottom: 10px; margin-top:-10px;">
#                 This tab will generate the same chart as the 
#                 <span style="background: #FF5E5E; padding:1px 4px; border-radius:6px; font-weight:600; font-size: 13px; color: black">Original Price</span> tab, 
#                 but the values are based on the converted amounts. 
#                 <br><br>
#                 Please note that the chart will 
#                 <span style="font-weight: bold; color: #FFE44D;">ONLY APPEAR</span> 
#                 if you use the currency converter feature. If you do not use it, the tab will display a message, as shown below.
#             </div>
#         """,
#         unsafe_allow_html=True
#     )

#     st.markdown(
#     """
#     <div style='background-color:#ffe6f2; padding:8px 12px; border-radius:8px; margin-bottom:15px;'>
#         <p style='font-size:13px; color:#a8326d; margin:4px;'>
#             💡 No converted data found. Please use the <b>Currency Converter</b> first.
#         </p>
#     </div>
#     """,
#     unsafe_allow_html=True
# )

# st.write("")
# st.markdown("**:violet-badge[6. COMPONENT COMPARISON]**")
# st.markdown(
#     """
#         <div style="text-align: justify; font-size: 15px; margin-bottom: 10px; margin-top:-10px;">
#             Similar to the rank visualization, this visualization also consists of two tabs: 
#             <span style="background: #FF5E5E; padding:1px 4px; border-radius:6px; font-weight:600; font-size: 13px; color: black">Original Price</span> and 
#             <span style="background: #FF00AA; padding:2px 4px; border-radius:6px; font-weight:600; font-size: 13px; color: black">Converted Price</span>. 
#             With this visualization, users can more easily identify the price differences of each component across vendors.
#         </div>
#     """,
#     unsafe_allow_html=True
# )

# tab1, tab2 = st.tabs(["Original Price", "Converted Price"])

# with tab1:
#     col1, col2 = st.columns(2)
#     with col1:
#         # st.image("assets/2.png")
#         st.write("tata")
#     with col2:
#         # st.image("assets/3.png")
#         st.write("tata")
# with tab2:
#     st.markdown(
#         """
#             <div style="text-align: justify; font-size: 15px; margin-bottom: 10px; margin-top:-10px;">
#                 This tab will generate the same chart as the 
#                 <span style="background: #FF5E5E; padding:1px 4px; border-radius:6px; font-weight:600; font-size: 13px; color: black">Original Price</span> tab, 
#                 but the values are based on the converted amounts. 
#                 <br><br>
#                 Please note that the chart will 
#                 <span style="font-weight: bold; color: #FFE44D;">ONLY APPEAR</span> if you use the currency converter feature. 
#                 If you do not use it, the tab will display a message, as shown below.
#             </div>
#         """,
#         unsafe_allow_html=True
#     )

#     st.markdown(
#     """
#     <div style='background-color:#ffe6f2; padding:8px 12px; border-radius:8px; margin-bottom:15px;'>
#         <p style='font-size:13px; color:#a8326d; margin:4px;'>
#             💡 No converted data found. Please use the <b>Currency Converter</b> first.
#         </p>
#     </div>
#     """,
#     unsafe_allow_html=True
# )
    
# st.write("")
# st.markdown("**:gray-badge[7. SUPER BUTTON]**")
# st.markdown(
#     """
#         <div style="text-align: justify; font-size: 15px; margin-bottom: 10px; margin-top:-10px;">
#             Lastly, there is a 
#             <span style="background:#FFCB09; padding:2px 4px; border-radius:6px; font-weight:600; font-size: 0.75rem; color: black">Super Button</span> 
#             feature where all dataframes generated by the system can be downloaded as a single file with multiple sheets. 
#             You can also customize the order of the sheets.
#             The interface looks more or less like this.
#         </div>
#     """,
#     unsafe_allow_html=True
# )

# dataframes = {
#     "Merge Data": df_merge,
#     "TCO Summary": df_tco,
#     "Bid & Price Analysis": df_analysis,
# }

# if "converted_tco_by_year" in st.session_state:
#     dataframes["TCO Converted"] = df_tco_converted

# # Tampilkan multiselect
# selected_sheets = st.multiselect(
#     "Select sheets to download in a single Excel file:",
#     options=list(dataframes.keys()),
#     default=list(dataframes.keys())  # default semua dipilih
# )

# # Fungsi "Super Button" & Formatting
# def generate_multi_sheet_excel(selected_sheets, df_dict):
#     """
#     Buat Excel multi-sheet dengan highlight:
#     - Sheet 'Bid & Price Analysis' -> highlight 1st & 2nd vendor
#     - Sheet lainnya -> highlight row TOTAL
#     """
#     output = BytesIO()

#     with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
#         for sheet in selected_sheets:
#             df = df_dict[sheet].copy()
#             df.to_excel(writer, index=False, sheet_name=sheet)
#             workbook  = writer.book
#             worksheet = writer.sheets[sheet]

#             # --- Format umum ---
#             fmt_rupiah = workbook.add_format({'num_format': '#,##0'})
#             fmt_pct    = workbook.add_format({'num_format': '#,##0.0"%"'})
#             fmt_total  = workbook.add_format({
#                 "bold": True, "bg_color": "#D9EAD3", "font_color": "#1A5E20", "num_format": "#,##0"
#             })
#             fmt_first  = workbook.add_format({'bg_color': '#C6EFCE', "num_format": "#,##0"})
#             fmt_second = workbook.add_format({'bg_color': '#FFEB9C', "num_format": "#,##0"})

#             # Identifikasi numeric columns
#             numeric_cols = df.select_dtypes(include=["number"]).columns.tolist()
#             vendor_cols = [c for c in numeric_cols] if sheet == "Bid & Price Analysis" else []

#             # Apply format kolom numeric / persen
#             for col_idx, col_name in enumerate(df.columns):
#                 if col_name in numeric_cols:
#                     worksheet.set_column(col_idx, col_idx, 15, fmt_rupiah)
#                 if "%" in col_name:
#                     worksheet.set_column(col_idx, col_idx, 15, fmt_pct)

#             # --- Highlight baris ---
#             for row_idx, row in enumerate(df.itertuples(index=False), start=1):
#                 # Cek apakah TOTAL
#                 is_total_row = any(str(x).strip().upper() == "TOTAL" for x in row if pd.notna(x))

#                 # Ambil nama 1st & 2nd vendor untuk sheet Bid & Price Analysis
#                 if sheet == "Bid & Price Analysis":
#                     first_vendor_name = row[df.columns.get_loc("1st Vendor")]
#                     second_vendor_name = row[df.columns.get_loc("2nd Vendor")]

#                     # Cari index kolom vendor di vendor_cols
#                     first_idx = df.columns.get_loc(first_vendor_name) if first_vendor_name in vendor_cols else None
#                     second_idx = df.columns.get_loc(second_vendor_name) if second_vendor_name in vendor_cols else None

#                 # Loop tiap kolom
#                 for col_idx, col_name in enumerate(df.columns):
#                     value = row[col_idx]
#                     fmt = None

#                     # Highlight TOTAL untuk sheet selain Bid & Price Analysis
#                     if is_total_row and sheet in ["Merge Data", "TCO Summary", "TCO Converted"]:
#                         fmt = fmt_total

#                     # Highlight 1st/2nd vendor
#                     elif sheet == "Bid & Price Analysis":
#                         if first_idx is not None and col_idx == first_idx:
#                             fmt = fmt_first
#                         elif second_idx is not None and col_idx == second_idx:
#                             fmt = fmt_second

#                     # Tangani NaN / None / inf
#                     if pd.isna(value) or (isinstance(value, (int, float)) and np.isinf(value)):
#                         value = ""

#                     worksheet.write(row_idx, col_idx, value, fmt)

#     output.seek(0)
#     return output

# # ---- DOWNLOAD BUTTON ----
# if selected_sheets:
#     excel_bytes = generate_multi_sheet_excel(selected_sheets, dataframes)

#     st.download_button(
#         label="Download",
#         data=excel_bytes,
#         file_name="Super Botton - TCO Comparison by Year.xlsx",
#         mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
#         type="primary",
#         use_container_width=True,
#     )

# st.write("")
# st.divider()

# st.markdown("#### Video Tutorial")
# st.markdown(
#     """
#         <div style="text-align: justify; font-size: 15px; margin-bottom: 10px; margin-top:-10px;">
#             I have also included a video tutorial, which you can access through the 
#             <span style="background:#FF0000; padding:2px 4px; border-radius:6px; font-weight:600; font-size: 0.75rem; color: black">YouTube</span> 
#             link below.
#         </div>
#     """,
#     unsafe_allow_html=True
# )

# st.video("https://youtu.be/_kqg84j2t-k?si=jpM6hcuqy5udC1Zc")