import streamlit as st
import pandas as pd
import io

# 1. 페이지 설정
st.set_page_config(page_title="AoN 시뮬레이터용 전처리 프로그램", layout="centered")

# ==========================================
# [기능 1] 로그인 시스템
# ==========================================
def login_system():
    if 'logged_in' not in st.session_state:
        st.session_state.logged_in = False

    if not st.session_state.logged_in:
        st.markdown(
            """
            <h3 style='text-align: center;'>🔒 로그인</h3>
            """, unsafe_allow_html=True
        )
        
        col1, col2, col3 = st.columns([1, 2, 1])
        with col2:
            input_id = st.text_input("ID", placeholder="아이디를 입력하세요")
            input_pw = st.text_input("Password", type="password", placeholder="비밀번호를 입력하세요")
            
            if st.button("로그인", type="primary", use_container_width=True):
                if input_id == "DMSERV" and input_pw == "ajr4ap8tjaga":
                    st.session_state.logged_in = True
                    st.success("로그인 성공!")
                    st.rerun()
                else:
                    st.error("ID 또는 비밀번호가 올바르지 않습니다.")
        st.stop()

# 로그인 체크 실행
login_system()

# ==========================================
# [기능 2] 메인 프로그램 시작
# ==========================================

# 디자인
st.markdown(
    """
    <h3 style='text-align: center;'>
        🏥<br>AoN 시뮬레이터용 전처리 프로그램
    </h3>
    """,
    unsafe_allow_html=True
)
st.markdown("---")

# 1. 파일 업로드
st.subheader("1️⃣ 데이터 파일 업로드")
st.info("📢 용량이 커서 오류가 발생하는 경우에, **CSV UTF-8**로 시트별로 저장 후 프로그램에 첨부해주세요.")

uploaded_file = st.file_uploader(
    "엑셀(.xlsx, .xls) 또는 CSV 파일을 선택하세요", 
    type=['xlsx', 'xls', 'csv']
)

if uploaded_file:
    try:
        df_raw = None
        selected_sheet = None
        
        file_ext = uploaded_file.name.split('.')[-1].lower()

        # [CASE A] 엑셀 파일 (.xlsx, .xls)
        if file_ext in ['xlsx', 'xls']:
            try:
                # calamine 엔진 시도
                try:
                    xls = pd.ExcelFile(uploaded_file, engine='calamine')
                except:
                    uploaded_file.seek(0)
                    xls = pd.ExcelFile(uploaded_file)
                
                sheet_names = xls.sheet_names
                
                st.subheader("2️⃣ 검사 항목(시트) 선택")
                selected_sheet = st.selectbox("데이터가 있는 시트를 선택하세요:", sheet_names)
                
                if selected_sheet:
                    df_raw = pd.read_excel(xls, sheet_name=selected_sheet)

            except Exception:
                # HTML 엑셀 시도
                try:
                    uploaded_file.seek(0)
                    html_dfs = pd.read_html(uploaded_file)
                    sheet_names = [f"Table_{i+1}" for i in range(len(html_dfs))]
                    st.session_state['html_dfs'] = html_dfs
                    
                    st.warning("⚠️ 웹 형식(HTML) 엑셀입니다.")
                    st.subheader("2️⃣ 검사 항목(시트) 선택")
                    selected_sheet = st.selectbox("데이터가 있는 시트를 선택하세요:", sheet_names)
                    
                    if selected_sheet:
                        idx = sheet_names.index(selected_sheet)
                        df_raw = st.session_state['html_dfs'][idx]
                except:
                    st.error("파일을 읽을 수 없습니다.")
                    st.stop()

        # [CASE B] CSV 파일
        elif file_ext == 'csv':
            try:
                df_raw = pd.read_csv(uploaded_file, encoding='utf-8')
            except UnicodeDecodeError:
                uploaded_file.seek(0)
                df_raw = pd.read_csv(uploaded_file, encoding='cp949')
            selected_sheet = "CSV_Data"

        # 데이터 로드 후 처리
        if df_raw is not None:
            columns = df_raw.columns.tolist()

            with st.expander("데이터 미리보기 (클릭)", expanded=False):
                st.dataframe(df_raw.head())

            st.markdown("---")
            
            # [기능 3] 스마트 검사명 인식 및 파일명 설정
            st.subheader("3️⃣ 파일명 설정 (자동 감지)")
            
            detected_test_name = None
            
            # 검사명 찾기
            target_col = None
            for col in columns:
                if any(k in str(col).lower() for k in ['검사명', '항목', 'test', 'item', 'name']):
                    target_col = col
                    break
            
            if target_col:
                try:
                    first_val = df_raw[target_col].dropna().iloc[0]
                    detected_test_name = str(first_val).strip()
                    st.success(f"🔍 '{target_col}' 열에서 검사명 '{detected_test_name}'을(를) 찾았습니다!")
                except:
                    pass
            
            if not detected_test_name:
                detected_test_name = str(selected_sheet)
                st.info(f"ℹ️ 검사명 열을 찾지 못해 시트명 '{selected_sheet}'을(를) 사용합니다.")

            col_hosp, col_test = st.columns(2)
            with col_hosp:
                hospital_name = st.text_input("병원명", placeholder="예: Severance")
            with col_test:
                final_test_name = st.text_input("검사명 (파일명 접미사)", value=detected_test_name)

            st.caption(f"예상 파일명: **{hospital_name if hospital_name else '[병원명]'} _ {final_test_name}.csv**")

            st.markdown("---")
            
            # 4. 컬럼 매칭
            st.subheader("4️⃣ 데이터 열(Column) 지정")
            col1, col2 = st.columns(2)
            
            default_date_idx = 0
            default_res_idx = 0
            
            lower_cols = [str(c).lower() for c in columns]
            
            for k in ['date', '일자', '일시', '시간', 'time', '시행']:
                for i, c in enumerate(lower_cols):
                    if k in c:
                        default_date_idx = i
                        break
                if default_date_idx != 0: break
            
            for k in ['result', '결과', '수치', 'val', 'measurement', 'data']:
                for i, c in enumerate(lower_cols):
                    if k in c:
                        default_res_idx = i
                        break
                if default_res_idx != 0: break
            
            with col1:
                date_col = st.selectbox("📅 날짜 열 (CSV B열)", columns, index=default_date_idx)
            with col2:
                meas_col = st.selectbox("📊 결과값 열 (CSV C열)", columns, index=default_res_idx)

            st.markdown("---")

            # -------------------------------------------------------------
            # [기능 4] 날짜 범위 제한 기능이 강화된 분할 옵션
            # -------------------------------------------------------------
            st.subheader("5️⃣ 데이터 분할 옵션")
            split_method = st.radio(
                "Training / Verification 분할 방식을 선택하세요:",
                ('8:2 비율 분할 (ID 기준)', '5:5 비율 분할 (ID 기준)', '날짜 기준 분할')
            )

            split_date = None
            
            if split_method == '날짜 기준 분할':
                # [핵심 로직] 선택된 날짜 컬럼을 미리 분석하여 min/max를 구함
                if date_col:
                    try:
                        # 날짜 변환 시도 (미리보기용)
                        temp_dates = pd.to_datetime(df_raw[date_col], errors='coerce').dropna()
                        
                        if not temp_dates.empty:
                            min_date = temp_dates.min().date()
                            max_date = temp_dates.max().date()
                            
                            st.write(f"📊 **데이터 날짜 범위:** {min_date} ~ {max_date}")
                            
                            # 달력 생성 (범위 제한 적용)
                            split_date = st.date_input(
                                "이 날짜까지 Training으로 설정합니다:",
                                value=min_date,      # 기본값: 시작 날짜
                                min_value=min_date,  # 최소 선택 가능 날짜
                                max_value=max_date   # 최대 선택 가능 날짜
                            )
                        else:
                            st.warning("⚠️ 선택한 날짜 열에서 유효한 날짜를 찾을 수 없습니다.")
                            split_date = st.date_input("날짜 선택")
                    except Exception as e:
                        st.warning(f"날짜 범위를 계산하는 중 오류가 발생했습니다: {e}")
                        split_date = st.date_input("날짜 선택")
            
            st.markdown("---")

            # 6. 변환 버튼
            if st.button("🚀 변환 및 CSV 생성", type="primary"):
                if not hospital_name:
                    st.warning("⚠️ 병원명을 입력해주세요!")
                else:
                    with st.spinner('처리 중...'):
                        try:
                            df_new = df_raw.loc[:, [date_col, meas_col]].copy()
                            df_new.columns = ['datetime', 'measurement']
                            
                            df_new = df_new.dropna(subset=['datetime', 'measurement'])
                            df_new['datetime'] = pd.to_datetime(df_new['datetime'], errors='coerce')
                            df_new = df_new.dropna(subset=['datetime'])
                            df_new = df_new.sort_values(by='datetime')
                            
                            df_new = df_new.reset_index(drop=True)
                            df_new['id'] = df_new.index + 1
                            
                            df_new['datetime'] = df_new['datetime'].dt.strftime('%Y-%m-%d')
                            df_new['measurement'] = pd.to_numeric(df_new['measurement'], errors='coerce').fillna(0).astype(int)

                            df_new['sheet_name'] = 'verification'
                            total_rows = len(df_new)

                            if split_method == '8:2 비율 분할 (ID 기준)':
                                split_idx = int(total_rows * 0.8)
                                df_new.loc[:split_idx-1, 'sheet_name'] = 'training'
                            elif split_method == '5:5 비율 분할 (ID 기준)':
                                split_idx = int(total_rows * 0.5)
                                df_new.loc[:split_idx-1, 'sheet_name'] = 'training'
                            elif split_method == '날짜 기준 분할' and split_date:
                                cutoff = split_date.strftime('%Y-%m-%d')
                                df_new.loc[df_new['datetime'] <= cutoff, 'sheet_name'] = 'training'

                            final_df = df_new[['id', 'datetime', 'measurement', 'sheet_name']]

                            csv_buffer = io.BytesIO()
                            final_df.to_csv(csv_buffer, index=False, encoding='utf-8-sig')
                            csv_data = csv_buffer.getvalue()

                            st.success(f"✅ 변환 완료! (총 {total_rows}개 데이터)")
                            
                            clean_hospital = hospital_name.strip().replace(" ", "_")
                            clean_test_name = str(final_test_name).strip().replace(" ", "_")
                            output_name = f"{clean_hospital}_{clean_test_name}.csv"

                            st.download_button(
                                label=f"📥 {output_name} 다운로드",
                                data=csv_data,
                                file_name=output_name,
                                mime="text/csv"
                            )
                            
                            st.write("▼ 결과 미리보기")
                            st.dataframe(final_df.head())

                        except Exception as e:
                            st.error(f"오류: {e}")

    except Exception as e:
        st.error(f"파일 오류: {e}")