import streamlit as st
import pandas as pd
from io import BytesIO
from openpyxl import Workbook
from openpyxl.styles import PatternFill, Font, Alignment

# ==============================
# 상단 UI
# ==============================

st.title("Crew Hotel Stay Generator")

uploaded_file = st.file_uploader("엑셀 파일 업로드", type=["xlsx"])

st.warning(
"""
⚠️ 숙박 계산 확인 필요

다음 경우 자동 계산이 부정확할 수 있습니다.

• 파일의 **첫 날짜**
• 파일의 **마지막 날짜**
• **마지막에서 두 번째 날짜**

이 경우 실제 예약 기록을 별도로 확인해주세요.
"""
)

# ==============================
# 파일 업로드
# ==============================

if uploaded_file:

    df = pd.read_excel(uploaded_file)

    df.columns = df.columns.str.strip()

    df = df[['Crew ID','Crew Name','Rank','Arr Flt','Arr Flt Date','Dep Flt','Dep Flt Date']]

    df['Arr Flt Date'] = pd.to_datetime(df['Arr Flt Date'], dayfirst=True)
    df['Dep Flt Date'] = pd.to_datetime(df['Dep Flt Date'], dayfirst=True)

    df = df.sort_values(['Crew ID','Arr Flt Date'])

    # ==============================
    # 연속 투숙 계산
    # ==============================

    stay_ids = []

    for crew_id, group in df.groupby('Crew ID'):

        group = group.sort_values('Arr Flt Date')

        stay_id = 0
        prev_dep = None

        for _, row in group.iterrows():

            arr = row['Arr Flt Date']
            dep = row['Dep Flt Date']

            if prev_dep is None:
                stay_id += 1

            elif (arr.date() - prev_dep.date()).days >= 1:
                stay_id += 1

            stay_ids.append(stay_id)

            prev_dep = dep

    df['stay_id'] = stay_ids

    # ==============================
    # 투숙 집계
    # ==============================

    stay_df = df.groupby(['Crew ID','Crew Name','Rank','stay_id']).agg(
        Checkin=('Arr Flt Date','min'),
        Checkout=('Dep Flt Date','max'),
        Arr_Flt=('Arr Flt','first'),
        Dep_Flt=('Dep Flt','last')
    ).reset_index()

    # ==============================
    # 박수 계산
    # ==============================

    stay_df['Nights_num'] = (
        stay_df['Checkout'].dt.normalize() -
        stay_df['Checkin'].dt.normalize()
    ).dt.days

    stay_df['Nights'] = stay_df['Nights_num'].astype(str) + "박"

    stay_df = stay_df[['Crew ID','Crew Name','Rank','Arr_Flt','Dep_Flt','Checkin','Checkout','Nights']]

    stay_df = stay_df.sort_values(['Checkin','Crew Name'])

    st.subheader("Crew Hotel Stay List")

    st.dataframe(stay_df, use_container_width=True)

    # ==============================
    # 엑셀 생성
    # ==============================

    output = BytesIO()

    wb = Workbook()
    ws = wb.active
    ws.title = "Stay List"

    header_fill = PatternFill(start_color="DDDDDD", end_color="DDDDDD", fill_type="solid")
    yellow_fill = PatternFill(start_color="FFFF00", end_color="FFFF00", fill_type="solid")

    prev_date = None

    for _, row in stay_df.iterrows():

        current_date = row['Checkin'].date()

        if prev_date != current_date:

            if prev_date is not None:
                ws.append([])

            ws.append([f"===== {current_date} ====="])

            cell = ws.cell(row=ws.max_row, column=1)
            cell.font = Font(bold=True, size=14)
            cell.alignment = Alignment(horizontal="center")

            ws.append(list(stay_df.columns))

            for col in range(1, len(stay_df.columns)+1):
                ws.cell(row=ws.max_row, column=col).fill = header_fill

        ws.append([
            row['Crew ID'],
            row['Crew Name'],
            row['Rank'],
            row['Arr_Flt'],
            row['Dep_Flt'],
            row['Checkin'].strftime("%Y-%m-%d"),
            row['Checkout'].strftime("%Y-%m-%d"),
            row['Nights']
        ])

        prev_date = current_date

    # ==============================
    # 열 폭 조정
    # ==============================

    column_widths = {
        "A":12,
        "B":20,
        "C":12,
        "D":12,
        "E":12,
        "F":15,
        "G":15,
        "H":8
    }

    for col, width in column_widths.items():
        ws.column_dimensions[col].width = width

    # ==============================
    # Crew ID 중복 색칠
    # ==============================

    duplicate_ids = stay_df['Crew ID'][stay_df['Crew ID'].duplicated(keep=False)]

    for row in ws.iter_rows(min_row=1, max_row=ws.max_row, min_col=1, max_col=1):

        cell = row[0]

        if cell.value in duplicate_ids.values:
            cell.fill = yellow_fill

    wb.save(output)
    output.seek(0)

    # ==============================
    # 다운로드
    # ==============================

    st.download_button(
        label="엑셀 다운로드",
        data=output,
        file_name="Crew_Hotel_Stay_List.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
