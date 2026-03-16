import streamlit as st
import pandas as pd
from io import BytesIO
import openpyxl 
from openpyxl import Workbook
from openpyxl.styles import PatternFill, Font, Alignment

# ==============================
# 상단 UI 및 안내 문구 (수정 및 보강)
# ==============================
st.set_page_config(page_title="Crew Stay Generator", layout="wide")
st.title("🏨 Crew Hotel Stay Generator")

# 주요 업데이트 및 오류 해결 안내
with st.expander("✅ 업데이트 내역 및 사용 가이드 (클릭하여 확인)", expanded=True):
    st.markdown("""
    ### 🚀 주요 업데이트 사항
    * **필터링 행 완전 제외**: 엑셀에서 **필터로 숨기거나 행 높이를 0으로 줄인 데이터**는 분석에서 완벽하게 제외됩니다.
    * **병합 셀 에러 해결**: 날짜 구분선(병합 셀) 생성 시 발생하던 열 너비 조정 오류를 수정했습니다.
    * **시간 정보 유지**: 엑셀 결과물에 체크인/아웃 시간이 정상적으로 표시됩니다.
    
    ### 🔍 오류 발생 시 체크리스트
    1.  **컬럼명 확인**: 업로드하는 엑셀에 `Crew ID`, `Crew Name`, `Rank`, `Arr Flt`, `Arr Flt Date`, `Dep Flt`, `Dep Flt Date` 컬럼이 정확히 포함되어 있는지 확인하세요.
    2.  **날짜 형식**: 날짜 컬럼에 텍스트나 깨진 값이 섞여 있으면 해당 행은 자동으로 제외됩니다.
    3.  **파일 닫기**: 엑셀 파일이 다른 프로그램에서 열려 있는 경우 업로드 오류가 발생할 수 있습니다.
    """)

st.info("💡 원본 엑셀에서 필요한 데이터만 필터링한 후 업로드하시면 더욱 정확합니다.")
uploaded_file = st.file_uploader("엑셀 파일 업로드", type=["xlsx"])

# ==============================
# 숨겨진 행을 제외하고 읽어오는 함수
# ==============================
def read_excel_visible_only(file):
    wb = openpyxl.load_workbook(file, data_only=True)
    ws = wb.active
    
    data = []
    # 첫 번째 행(헤더) 가져오기
    headers = [cell.value for cell in ws[1]]
    
    # 두 번째 행부터 순회하며 숨겨진 행인지 체크
    for row in ws.iter_rows(min_row=2):
        row_idx = row[0].row
        # 숨김 상태이거나 행 높이가 0인 경우 제외
        if ws.row_dimensions[row_idx].hidden == False and ws.row_dimensions[row_idx].height != 0:
            data.append([cell.value for cell in row])
            
    return pd.DataFrame(data, columns=headers)

# ==============================
# 파일 처리 시작 (기존 로직 유지)
# ==============================
if uploaded_file:
    df = read_excel_visible_only(uploaded_file)
    
    if df.empty:
        st.error("데이터가 없거나 모든 행이 숨겨져 있습니다.")
        st.stop()

    df.columns = df.columns.str.strip()
    
    required_cols = ['Crew ID','Crew Name','Rank','Arr Flt','Arr Flt Date','Dep Flt','Dep Flt Date']
    df = df[required_cols].copy()

    df['Arr Flt Date'] = pd.to_datetime(df['Arr Flt Date'], dayfirst=True, errors='coerce')
    df['Dep Flt Date'] = pd.to_datetime(df['Dep Flt Date'], dayfirst=True, errors='coerce')
    df = df.dropna(subset=['Crew ID', 'Arr Flt Date', 'Dep Flt Date'])

    df = df.sort_values(['Crew ID','Arr Flt Date'])

    # 연속 투숙 계산
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

            stay_ids.append(f"{crew_id}_{stay_id}")
            prev_dep = dep

    df['stay_id'] = stay_ids

    stay_df = df.groupby(['Crew ID','Crew Name','Rank','stay_id']).agg(
        Checkin=('Arr Flt Date','min'),
        Checkout=('Dep Flt Date','max'),
        Arr_Flt=('Arr Flt','first'),
        Dep_Flt=('Dep Flt','last')
    ).reset_index()

    stay_df['Nights_num'] = (stay_df['Checkout'].dt.date - stay_df['Checkin'].dt.date).apply(lambda x: x.days)
    stay_df['Nights'] = stay_df['Nights_num'].astype(str) + "박"
    
    stay_df = stay_df[['Crew ID','Crew Name','Rank','Arr_Flt','Dep_Flt','Checkin','Checkout','Nights']]
    stay_df = stay_df.sort_values(['Checkin','Crew Name'])

    st.subheader("🗓️ Crew Hotel Stay List (필터링 완료)")
    st.dataframe(stay_df, use_container_width=True)

    # 엑셀 생성
    output = BytesIO()
    wb_out = Workbook()
    ws_out = wb_out.active
    ws_out.title = "Stay List"

    header_fill = PatternFill(start_color="DDDDDD", end_color="DDDDDD", fill_type="solid")
    yellow_fill = PatternFill(start_color="FFFF00", end_color="FFFF00", fill_type="solid")

    prev_date = None
    duplicate_ids = stay_df['Crew ID'].astype(str)[stay_df['Crew ID'].astype(str).duplicated(keep=False)].unique()

    for _, row in stay_df.iterrows():
        current_date = row['Checkin'].date()

        if prev_date != current_date:
            if prev_date is not None:
                ws_out.append([])

            ws_out.append([f"===== {current_date} ====="])
            cell = ws_out.cell(row=ws_out.max_row, column=1)
            cell.font = Font(bold=True, size=14)
            cell.alignment = Alignment(horizontal="center")
            ws_out.merge_cells(start_row=ws_out.max_row, start_column=1, end_row=ws_out.max_row, end_column=8)

            ws_out.append(list(stay_df.columns))
            for col in range(1, 9):
                ws_out.cell(row=ws_out.max_row, column=col).fill = header_fill

        ws_out.append([
            str(row['Crew ID']), row['Crew Name'], row['Rank'],
            row['Arr_Flt'], row['Dep_Flt'],
            row['Checkin'].strftime("%Y-%m-%d %H:%M"),
            row['Checkout'].strftime("%Y-%m-%d %H:%M"),
            row['Nights']
        ])

        if str(row['Crew ID']) in duplicate_ids:
            ws_out.cell(row=ws_out.max_row, column=1).fill = yellow_fill

        prev_date = current_date

    column_widths = {"A":15, "B":20, "C":12, "D":12, "E":12, "F":20, "G":20, "H":8}
    for col, width in column_widths.items():
        ws_out.column_dimensions[col].width = width

    wb_out.save(output)
    output.seek(0)

    st.download_button(
        label="📥 최종 엑셀 다운로드",
        data=output,
        file_name="Crew_Reservation_Final.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
