import streamlit as st
import pandas as pd
from io import BytesIO
import os
import json

# --- 설정 값 ---
DB_TARGET_COL = "품목코드"  # DB의 B열(코드) 머리글 이름
BOM_TARGET_COL = "ERP CODE" # BOM의 M열(코드) 머리글 이름
CONFIG_FILE = "db_config.json"

st.set_page_config(page_title="BOM 자동 검증", layout="wide")

st.title("📊 BOM vs Database 자동 체크 (1차 & 2차 공통 참조)")
st.write("DB의 [코드, SPEC, PN] 마스터를 바탕으로 1차 및 2차 자재 교차 검증")

# --- 설정 파일(JSON) 읽기/쓰기 함수 ---
def load_db_path():
    if os.path.exists(CONFIG_FILE):
        with open(CONFIG_FILE, 'r', encoding='utf-8') as f:
            config = json.load(f)
            return config.get('db_path', '')
    return ''

def save_db_path(path):
    with open(CONFIG_FILE, 'w', encoding='utf-8') as f:
        json.dump({'db_path': path}, f)

# --- 텍스트 정제 함수 ---
def clean_for_compare(val):
    if pd.isna(val):
        return ""
    val_str = str(val).strip()
    if val_str.endswith('.0'):
        val_str = val_str[:-2]
    return val_str.upper().replace(" ", "").replace("\n", "").replace("\r", "")

# --- 1️⃣ DB 파일 설정 ---
st.subheader("⚙️ 마스터 DB 연결")
saved_db_path = load_db_path()
df_db = None

col_db1, col_db2 = st.columns([3, 1])

with col_db1:
    if saved_db_path and os.path.exists(saved_db_path):
        st.success(f"🔗 현재 연결된 DB: `{saved_db_path}`")
        db_row_input = st.text_input("DB 품목코드 위치 (행 번호)", value="8", key="db_row")
        
        try:
            db_row = int(db_row_input)
            if saved_db_path.endswith('.csv'):
                df_db = pd.read_csv(saved_db_path, header=db_row-1, encoding='utf-8-sig', on_bad_lines='skip')
            else:
                df_db = pd.read_excel(saved_db_path, header=db_row-1)
            
            # DB 컬럼명 공백만 제거 (Unnamed는 그대로 유지)
            df_db.columns = df_db.columns.astype(str).str.strip()
            
        except Exception as e:
            st.error(f"DB 파일을 읽는 중 오류가 발생했습니다: {e}")
            
    else:
        st.info("💡 처음 한 번만 DB 파일의 경로를 입력해 주세요. (예: C:/Users/user/Desktop/DB.xlsx)")
        new_path = st.text_input("DB 파일 절대 경로 입력:")
        if st.button("경로 저장"):
            clean_path = new_path.strip().strip("'").strip('"') 
            if os.path.exists(clean_path):
                save_db_path(clean_path)
                st.success("경로가 저장되었습니다! 앱을 새로고침합니다.")
                st.rerun()
            else:
                st.error("⚠️ 해당 경로에 파일이 존재하지 않습니다. 경로와 확장자를 다시 확인해 주세요.")

with col_db2:
    if saved_db_path:
        if st.button("🔄 DB 경로 변경/초기화"):
            save_db_path("")
            st.rerun()

st.divider()

# --- 2️⃣ BOM 파일 업로드 ---
st.subheader("📁 검증할 BOM 파일 업로드")
col1, col2 = st.columns(2)
with col1:
    bom_file = st.file_uploader("BOM 엑셀/CSV 업로드", type=['xlsx', 'csv'])
with col2:
    bom_row_input = st.text_input("BOM ERP CODE 위치 (행 번호)", value="13", key="bom_row")

# --- 검증 로직 ---
if df_db is not None and bom_file:
    try:
        if not bom_row_input.isdigit():
            st.error("⚠️ 행 번호는 숫자로만 입력해 주세요.")
            st.stop()
            
        bom_row = int(bom_row_input)

        if bom_file.name.endswith('.csv'):
            df_bom = pd.read_csv(bom_file, header=bom_row-1, encoding='utf-8-sig', on_bad_lines='skip')
        else:
            df_bom = pd.read_excel(bom_file, header=bom_row-1)

        # BOM 컬럼명 공백만 제거 (Unnamed는 그대로 유지)
        df_bom.columns = df_bom.columns.astype(str).str.strip()

        # 원인 파악용 에러 메시지 보강
        if DB_TARGET_COL not in df_db.columns:
            st.error(f"⚠️ DB에서 '{DB_TARGET_COL}' 열을 찾을 수 없습니다. (현재 행 번호: {db_row})")
            st.info(f"🔍 DB 컬럼 목록: {df_db.columns.tolist()}")
            st.stop()
            
        if BOM_TARGET_COL not in df_bom.columns:
            st.error(f"⚠️ BOM에서 '{BOM_TARGET_COL}' 열을 찾을 수 없습니다. (현재 행 번호: {bom_row})")
            st.info(f"🔍 BOM 컬럼 목록: {df_bom.columns.tolist()}")
            st.stop()

        if st.button("🚀 비교 시작"):
            col_db_spec_name = df_db.columns[3] # D열 (SPEC)
            col_db_pn_name = df_db.columns[4]   # E열 (P/N)
            
            db_mapping = df_db[[DB_TARGET_COL, col_db_spec_name, col_db_pn_name]].copy()
            db_mapping = db_mapping.drop_duplicates(subset=[DB_TARGET_COL])
            db_mapping.columns = ['DB_KEY', 'DB_SPEC_VAL', 'DB_PN_VAL']

            # 💡 BOM 컬럼 위치 지정 (알려주신 위치로 업데이트 완료!)
            col_bom_spec_name = df_bom.columns[3]   # D열 (1차/2차 공통 SPEC)
            col_bom_pn_name = df_bom.columns[12]    # M열 (1차 P/N)
            col_bom_tier2_code = df_bom.columns[13] # N열 (2차 업체 코드)
            col_bom_tier2_pn = df_bom.columns[15]   # P열 (2차 업체 P/N)

            result_merged = df_bom.merge(db_mapping, left_on=BOM_TARGET_COL, right_on='DB_KEY', how='left')
            result_merged.rename(columns={'DB_KEY': 'DB_KEY_1', 'DB_SPEC_VAL': 'DB_SPEC_1', 'DB_PN_VAL': 'DB_PN_1'}, inplace=True)

            result_merged = result_merged.merge(db_mapping, left_on=col_bom_tier2_code, right_on='DB_KEY', how='left')
            result_merged.rename(columns={'DB_KEY': 'DB_KEY_2', 'DB_SPEC_VAL': 'DB_SPEC_2', 'DB_PN_VAL': 'DB_PN_2'}, inplace=True)

            def do_validation(row):
                # --- 1차 검증 ---
                res_1st = "정상"
                ref_1st_str = ""
                
                if pd.isna(row['DB_KEY_1']) or str(row['DB_KEY_1']).strip() == "":
                    res_1st = "코드 없음"
                else:
                    bom_spec_cmp = clean_for_compare(row[col_bom_spec_name])
                    bom_pn_cmp = clean_for_compare(row[col_bom_pn_name])
                    db_spec_cmp = clean_for_compare(row['DB_SPEC_1'])
                    db_pn_cmp = clean_for_compare(row['DB_PN_1'])
                    
                    err_1st = []
                    ref_1st = []
                    if bom_spec_cmp != db_spec_cmp:
                        err_1st.append("SPEC 오류")
                        ref_1st.append(f"SPEC: {row['DB_SPEC_1'] if not pd.isna(row['DB_SPEC_1']) else ''}")
                    if bom_pn_cmp != db_pn_cmp:
                        err_1st.append("PN 오류")
                        ref_1st.append(f"PN: {row['DB_PN_1'] if not pd.isna(row['DB_PN_1']) else ''}")
                    
                    if err_1st:
                        res_1st = ", ".join(err_1st)
                        ref_1st_str = ", ".join(ref_1st)

                # --- 2차 검증 ---
                res_2nd = ""
                ref_2nd_str = ""
                
                bom_t2_code_val = row[col_bom_tier2_code]
                
                if pd.notna(bom_t2_code_val) and str(bom_t2_code_val).strip() != "":
                    if pd.isna(row['DB_KEY_2']) or str(row['DB_KEY_2']).strip() == "":
                        res_2nd = "DB에 코드 없음" 
                    else:
                        # 💡 핵심: 2차 품목도 D열(col_bom_spec_name)을 가져와서 검사하도록 수정!
                        bom_t2_spec_cmp = clean_for_compare(row[col_bom_spec_name])
                        db_t2_spec_cmp = clean_for_compare(row['DB_SPEC_2'])
                        
                        bom_t2_pn_cmp = clean_for_compare(row[col_bom_tier2_pn])
                        db_t2_pn_cmp = clean_for_compare(row['DB_PN_2'])
                        
                        err_2nd = []
                        ref_2nd = []
                        
                        if bom_t2_spec_cmp != db_t2_spec_cmp:
                            err_2nd.append("2차SPEC 오류")
                            ref_2nd.append(f"2차SPEC: {row['DB_SPEC_2'] if not pd.isna(row['DB_SPEC_2']) else ''}")
                            
                        if bom_t2_pn_cmp != db_t2_pn_cmp:
                            err_2nd.append("2차PN 오류")
                            ref_2nd.append(f"2차PN: {row['DB_PN_2'] if not pd.isna(row['DB_PN_2']) else ''}")
                            
                        if err_2nd:
                            res_2nd = ", ".join(err_2nd)
                            ref_2nd_str = ", ".join(ref_2nd)
                        else:
                            res_2nd = "정상"
                            
                return pd.Series([res_1st, ref_1st_str, res_2nd, ref_2nd_str])

            result_merged[['검증결과', '참조_DB데이터', '검증결과(2차)', '참조_DB데이터 (2차)']] = result_merged.apply(do_validation, axis=1)

            # 불필요한 임시 열 제거 (Unnamed 열은 제거하지 않고 남겨둡니다)
            drop_cols = ['DB_KEY_1', 'DB_SPEC_1', 'DB_PN_1', 'DB_KEY_2', 'DB_SPEC_2', 'DB_PN_2']
            result = result_merged.drop(columns=[c for c in drop_cols if c in result_merged.columns])

            # 컬럼 순서 조정
            new_cols = ['검증결과', '참조_DB데이터', '검증결과(2차)', '참조_DB데이터 (2차)']
            cols = new_cols + [c for c in result.columns if c not in new_cols]
            result = result[cols]

            # --- 색상 스타일링 ---
            def highlight_result(val):
                val_str = str(val)
                if val_str == "정상":
                    return 'background-color: #e6ffed; color: #1a7f37; font-weight: bold;'
                elif "코드 없음" in val_str:
                    return 'background-color: #f6f8fa; color: #57606a;'
                elif "오류" in val_str:
                    if "," in val_str: 
                        return 'background-color: #ffebe9; color: #cf222e; font-weight: bold;'
                    elif "PN" in val_str: 
                        return 'background-color: #f3e8ff; color: #6e32c9; font-weight: bold;'
                    else: 
                        return 'background-color: #fff8c5; color: #9a6700; font-weight: bold;'
                return ''

            st.success("비교 완료! 1차 및 2차 품목 모두 SPEC과 P/N 검증이 완료되었습니다.")
            
            target_styled_cols = ['검증결과', '검증결과(2차)']
            try:
                styled_result = result.style.map(highlight_result, subset=target_styled_cols)
            except AttributeError:
                styled_result = result.style.applymap(highlight_result, subset=target_styled_cols)

            st.dataframe(styled_result, use_container_width=True)

            output = BytesIO()
            with pd.ExcelWriter(output, engine='openpyxl') as writer:
                styled_result.to_excel(writer, index=False, sheet_name="검증결과")
            
            st.download_button("📥 엑셀 다운로드 (색상 포함)", output.getvalue(), "BOM_검증결과.xlsx")

    except Exception as e:
        st.error(f"오류: {e}")

elif df_db is None:
    st.info("👆 위에서 마스터 DB 경로를 설정하고 정상적으로 연결되면 BOM 업로드 창이 나타납니다.")
else:
    st.info("👇 위에서 비교할 BOM 파일을 업로드해 주세요.")