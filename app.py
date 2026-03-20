import streamlit as st
import pandas as pd
from io import BytesIO

# --- 설정 값 ---
DB_TARGET_COL = "품목코드"
BOM_TARGET_COL = "ERP CODE"

st.set_page_config(page_title="BOM 자동 검증", layout="wide")

st.title("📊 BOM vs Database 자동 체크 (1차 & 2차 공통 참조)")
st.write("마스터 DB와 BOM 파일을 업로드하여 1차 및 2차 자재 교차 검증")

# --- 텍스트 정제 함수 ---
def clean_for_compare(val):
    if pd.isna(val):
        return ""
    val_str = str(val).strip()
    if val_str.endswith('.0'):
        val_str = val_str[:-2]
    return val_str.upper().replace(" ", "").replace("\n", "").replace("\r", "")

# --- 1️⃣ 마스터 DB 파일 업로드 (가장 중요한 수정 부분!) ---
st.subheader("⚙️ 마스터 DB 연결")
df_db = None

# 로컬 경로 입력 대신 파일 업로더 사용
db_file = st.file_uploader("1️⃣ 마스터 DB 엑셀 업로드", type=['xlsx', 'csv'])
db_row_input = st.text_input("DB 품목코드 위치 (행 번호)", value="8")

if db_file:
    try:
        db_row = int(db_row_input)
        if db_file.name.endswith('.csv'):
            df_db = pd.read_csv(db_file, header=db_row-1, encoding='utf-8-sig', on_bad_lines='skip')
        else:
            df_db = pd.read_excel(db_file, header=db_row-1)
        
        df_db.columns = df_db.columns.astype(str).str.strip()
        st.success("🔗 마스터 DB 로드 완료!")
        
    except Exception as e:
        st.error(f"DB 파일을 읽는 중 오류가 발생했습니다: {e}")

st.divider()

# --- 2️⃣ BOM 파일 업로드 ---
st.subheader("📁 검증할 BOM 파일 업로드")
df_bom = None
bom_file = st.file_uploader("2️⃣ 비교할 BOM 엑셀 업로드", type=['xlsx', 'csv'])
bom_row_input = st.text_input("BOM ERP CODE 위치 (행 번호)", value="13")

if bom_file:
    try:
        bom_row = int(bom_row_input)
        if bom_file.name.endswith('.csv'):
            df_bom = pd.read_csv(bom_file, header=bom_row-1, encoding='utf-8-sig', on_bad_lines='skip')
        else:
            df_bom = pd.read_excel(bom_file, header=bom_row-1)
            
        df_bom.columns = df_bom.columns.astype(str).str.strip()
        st.success("📁 BOM 파일 로드 완료!")
        
    except Exception as e:
        st.error(f"BOM 파일을 읽는 중 오류가 발생했습니다: {e}")

# --- 검증 및 결과 처리 ---
if df_db is not None and df_bom is not None:
    # 컬럼 존재 유무 확인
    if DB_TARGET_COL not in df_db.columns or BOM_TARGET_COL not in df_bom.columns:
        st.error(f"⚠️ 컬럼명을 찾을 수 없습니다. DB에는 '{DB_TARGET_COL}', BOM에는 '{BOM_TARGET_COL}' 컬럼이 있어야 합니다.")
    else:
        if st.button("🚀 비교 시작"):
            # DB 기준 테이블 생성 (B열:품목코드, D열:SPEC, E열:PN)
            col_db_spec_name = df_db.columns[3] # D열 (SPEC)
            col_db_pn_name = df_db.columns[4]   # E열 (P/N)
            
            db_mapping = df_db[[DB_TARGET_COL, col_db_spec_name, col_db_pn_name]].copy()
            db_mapping = db_mapping.drop_duplicates(subset=[DB_TARGET_COL])
            db_mapping.columns = ['DB_KEY', 'DB_SPEC_VAL', 'DB_PN_VAL']

            # BOM 컬럼 위치 지정 (알려주신 위치 적용)
            col_bom_spec_name = df_bom.columns[3]   # D열 (1차 SPEC)
            col_bom_pn_name = df_bom.columns[12]    # M열 (1차 P/N)
            col_bom_tier2_code = df_bom.columns[13] # N열 (2차 업체 코드)
            col_bom_tier2_pn = df_bom.columns[15]   # P열 (2차 업체 P/N)

            # 1차 병합
            result_merged = df_bom.merge(db_mapping, left_on=BOM_TARGET_COL, right_on='DB_KEY', how='left')
            result_merged.rename(columns={'DB_KEY': 'DB_KEY_1', 'DB_SPEC_VAL': 'DB_SPEC_1', 'DB_PN_VAL': 'DB_PN_1'}, inplace=True)

            # 2차 병합
            result_merged = result_merged.merge(db_mapping, left_on=col_bom_tier2_code, right_on='DB_KEY', how='left')
            result_merged.rename(columns={'DB_KEY': 'DB_KEY_2', 'DB_SPEC_VAL': 'DB_SPEC_2', 'DB_PN_VAL': 'DB_PN_2'}, inplace=True)

            # 검증 함수
            def do_validation(row):
                # 1차 검증
                res_1st = "정상"
                ref_1st_str = ""
                if pd.isna(row['DB_KEY_1']):
                    res_1st = "코드 없음"
                else:
                    if clean_for_compare(row[col_bom_spec_name]) != clean_for_compare(row['DB_SPEC_1']):
                        res_1st = "SPEC 오류"
                        ref_1st_str = f"SPEC: {row['DB_SPEC_1']}"
                    elif clean_for_compare(row[col_bom_pn_name]) != clean_for_compare(row['DB_PN_1']):
                        res_1st = "PN 오류"
                        ref_1st_str = f"PN: {row['DB_PN_1']}"
                
                # 2차 검증 (N열에 코드가 있을 때만)
                res_2nd = ""
                ref_2nd_str = ""
                bom_t2_code_val = row[col_bom_tier2_code]
                if pd.notna(bom_t2_code_val) and str(bom_t2_code_val).strip() != "":
                    if pd.isna(row['DB_KEY_2']):
                        res_2nd = "DB에 코드 없음"
                    else:
                        if clean_for_compare(row[col_bom_tier2_pn]) != clean_for_compare(row['DB_PN_2']):
                            res_2nd = "2차PN 오류"
                            ref_2nd_str = f"2차PN: {row['DB_PN_2']}"
                        else:
                            res_2nd = "정상"
                            
                return pd.Series([res_1st, ref_1st_str, res_2nd, ref_2nd_str])

            result_merged[['검증결과', '참조_DB데이터', '검증결과(2차)', '참조_DB데이터 (2차)']] = result_merged.apply(do_validation, axis=1)

            # 결과 정리
            drop_cols = ['DB_KEY_1', 'DB_SPEC_1', 'DB_PN_1', 'DB_KEY_2', 'DB_SPEC_2', 'DB_PN_2']
            result = result_merged.drop(columns=[c for c in drop_cols if c in result_merged.columns])
            new_cols = ['검증결과', '참조_DB데이터', '검증결과(2차)', '참조_DB데이터 (2차)']
            cols = new_cols + [c for c in result.columns if c not in new_cols]
            result = result[cols]

            # 스타일링 및 출력
            def highlight_result(val):
                if val == "정상": return 'background-color: #e6ffed; color: #1a7f37; font-weight: bold;'
                elif "오류" in str(val): return 'background-color: #ffebe9; color: #cf222e; font-weight: bold;'
                elif "코드 없음" in str(val): return 'background-color: #f6f8fa; color: #57606a;'
                return ''

            styled_result = result.style.map(highlight_result, subset=['검증결과', '검증결과(2차)'])
            st.dataframe(styled_result, use_container_width=True)

            # 다운로드 버튼
            output = BytesIO()
            with pd.ExcelWriter(output, engine='openpyxl') as writer:
                styled_result.to_excel(writer, index=False, sheet_name="검증결과")
            st.download_button("📥 엑셀 다운로드 (색상 포함)", output.getvalue(), "BOM_검증결과.xlsx")

elif db_file is None:
    st.info("👆 먼저 1차 마스터 DB 파일을 업로드해 주세요.")
elif bom_file is None:
    st.info("👆 비교할 BOM 파일을 업로드해 주세요.")