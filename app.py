import streamlit as st
import pandas as pd
import os
import plotly.graph_objects as go

# 페이지 설정
st.set_page_config(
    page_title="토마토 적정생육표 매칭 시스템",
    page_icon="🍅",
    layout="wide"
)

st.title("🍅 토마토 적정생육표 매칭 시스템")
st.markdown("---")

# 데이터 파일 경로
DATA_DIR = "data"

# Excel 파일 목록
EXCEL_FILES = {
    "일사량별": {
        "비닐": "일사량별_비닐_적정생육표.xlsx",
        "유리": "일사량별_유리_적정생육표.xlsx"
    },
    "생육상태별": {
        "비닐": "생육상태별_비닐_적정생육표.xlsx",
        "유리": "생육상태별_유리_적정생육표.xlsx"
    }
}



@st.cache_data
def load_excel_sheets(file_path: str) -> dict[str, pd.DataFrame]:
    """Excel 파일의 모든 시트를 로드"""
    try:
        excel_file = pd.ExcelFile(file_path)
        sheets = {}
        for sheet_name in excel_file.sheet_names:
            sheets[sheet_name] = pd.read_excel(file_path, sheet_name=sheet_name)
        return sheets
    except Exception as e:
        st.error(f"Excel 파일 읽기 오류: {e}")
        return {}

@st.cache_data
def get_sheet_names(file_path: str) -> list[str]:
    """Excel 파일의 시트명 목록 반환"""
    try:
        excel_file = pd.ExcelFile(file_path)
        return excel_file.sheet_names
    except Exception as e:
        st.error(f"시트명 읽기 오류: {e}")
        return []

def validate_user_data(df: pd.DataFrame) -> tuple:
    """
    사용자 데이터에서 사용 가능한 컬럼을 찾기
    실제 제공된 정확한 컬럼명만 유효
    """
    # 실제 Excel 파일의 정확한 컬럼명 매핑
    exact_columns_map = {
        '누적일사량': ['누적일사량(범위)'],
        '외기기온': ['외기기온(범위)'],
        '생산량': ['생산량(㎏/3.3㎡)'],
        '일일평균온도': ['일일 평균온도(℃)'],
        '주간평균온도': ['주간 평균온도(℃)'],
        '야간평균온도': ['야간 평균온도(℃)'],
        '새벽온도': ['새벽온도(℃)'],
        '주간평균습도': ['주간 평균습도(%)'],
        '잔존CO2': ['잔존 CO₂(ppm)'],
        '급액EC': ['급액 EC(dS/m)'],
        '급액pH': ['급액 pH'],
        '1회급액량': ['1회 급액량(㏄/회)'],
        '1일공급량': ['1일 공급량(㏄/day)']
    }
    
    found_columns = {}
    missing_columns = []
    
    # 각 컬럼에 대해 정확한 컬럼명을 찾기 (있는 것만)
    for key, possible_names in exact_columns_map.items():
        found = False
        for possible_name in possible_names:
            if possible_name in df.columns:
                found_columns[key] = possible_name
                found = True
                break
        
        if not found:
            missing_columns.append(key)
    
    # 최소 1개 컬럼이라도 있으면 유효한 것으로 처리
    is_valid = len(found_columns) > 0
    return is_valid, missing_columns, found_columns

def find_matching_range(reference_df: pd.DataFrame, user_data: pd.DataFrame, column_mapping: dict) -> tuple:
    """
    사용자 데이터의 누적일사량/외기기온 범위를 모두 포함하는 적정생육표 구간을 찾기
    """
    # 컬럼 매핑에서 실제 컬럼명 가져오기
    radiation_col = column_mapping.get('누적일사량')
    temp_col = column_mapping.get('외기기온')
    
    # 적정생육표의 해당 컬럼 확인 (정확한 컬럼명)
    ref_radiation_col = None
    ref_temp_col = None
    
    # 적정생육표에서 누적일사량 컬럼 찾기 (정확한 컬럼명만)
    if '누적일사량(범위)' in reference_df.columns:
        ref_radiation_col = '누적일사량(범위)'
    
    # 적정생육표에서 외기기온 컬럼 찾기 (정확한 컬럼명만)
    if '외기기온(범위)' in reference_df.columns:
        ref_temp_col = '외기기온(범위)'
    
    # 사용자 데이터의 범위 계산
    user_radiation_min = None
    user_radiation_max = None
    user_temp_min = None
    user_temp_max = None
    
    # 사용자 데이터에서 누적일사량 범위 계산
    if radiation_col and radiation_col in user_data.columns:
        try:
            user_radiation_data = pd.to_numeric(user_data[radiation_col], errors='coerce').dropna()
            if len(user_radiation_data) > 0:
                user_radiation_min = user_radiation_data.min()
                user_radiation_max = user_radiation_data.max()
                st.info(f"사용자 누적일사량 범위: {user_radiation_min:.2f} ~ {user_radiation_max:.2f}")
        except Exception as e:
            st.warning(f"누적일사량 데이터 처리 중 오류: {e}")
    
    # 사용자 데이터에서 외기기온 범위 계산
    if temp_col and temp_col in user_data.columns:
        try:
            user_temp_data = pd.to_numeric(user_data[temp_col], errors='coerce').dropna()
            if len(user_temp_data) > 0:
                user_temp_min = user_temp_data.min()
                user_temp_max = user_temp_data.max()
                st.info(f"사용자 외기기온 범위: {user_temp_min:.2f} ~ {user_temp_max:.2f}")
        except Exception as e:
            st.warning(f"외기기온 데이터 처리 중 오류: {e}")
    
    # 적정생육표에서 사용자 범위를 포함하는 구간 찾기
    matched_rows = []
    
    for idx, row in reference_df.iterrows():
        is_match = True
        
        # 누적일사량 범위 체크
        if ref_radiation_col and user_radiation_min is not None and user_radiation_max is not None:
            try:
                ref_radiation_str = str(row[ref_radiation_col])
                # 범위 형태 파싱 (예: "10-15", "10~15", "10 - 15" 등)
                if '-' in ref_radiation_str or '~' in ref_radiation_str:
                    # 구분자 찾기
                    separator = '-' if '-' in ref_radiation_str else '~'
                    parts = ref_radiation_str.split(separator)
                    if len(parts) == 2:
                        ref_min = float(parts[0].strip())
                        ref_max = float(parts[1].strip())
                        # 적정생육표 범위가 사용자 범위를 포함하는지 확인
                        if not (ref_min <= user_radiation_min and ref_max >= user_radiation_max):
                            is_match = False
                else:
                    # 단일 값인 경우
                    ref_val = float(ref_radiation_str)
                    if not (user_radiation_min <= ref_val <= user_radiation_max):
                        is_match = False
            except (ValueError, TypeError):
                # 파싱 실패 시 해당 행은 제외
                is_match = False
        
        # 외기기온 범위 체크
        if ref_temp_col and user_temp_min is not None and user_temp_max is not None and is_match:
            try:
                ref_temp_str = str(row[ref_temp_col])
                # 범위 형태 파싱
                if '-' in ref_temp_str or '~' in ref_temp_str:
                    separator = '-' if '-' in ref_temp_str else '~'
                    parts = ref_temp_str.split(separator)
                    if len(parts) == 2:
                        ref_min = float(parts[0].strip())
                        ref_max = float(parts[1].strip())
                        # 적정생육표 범위가 사용자 범위를 포함하는지 확인
                        if not (ref_min <= user_temp_min and ref_max >= user_temp_max):
                            is_match = False
                else:
                    # 단일 값인 경우
                    ref_val = float(ref_temp_str)
                    if not (user_temp_min <= ref_val <= user_temp_max):
                        is_match = False
            except (ValueError, TypeError):
                is_match = False
        
        if is_match:
            matched_rows.append(row)
    
    if matched_rows:
        matched_reference = pd.DataFrame(matched_rows)
    else:
        matched_reference = pd.DataFrame()
    
    return matched_reference, user_data

def create_comparison_charts(reference_df: pd.DataFrame, user_df: pd.DataFrame, column_mapping: dict):
    """
    게이지 차트로 사용자 평균값과 적정생육표 범위 비교
    """
    # 모든 비교 항목들
    all_comparisons = [
        ('일일 평균온도(℃)', '일일평균온도', '일일평균온도', '°C'),
        ('주간 평균온도(℃)', '주간평균온도', '주간평균온도', '°C'),
        ('야간 평균온도(℃)', '야간평균온도', '야간평균온도', '°C'),
        ('새벽온도(℃)', '새벽온도', '새벽온도', '°C'),
        ('주간 평균습도(%)', '주간평균습도', '주간평균습도', '%'),
        ('생산량(㎏/3.3㎡)', '생산량', '생산량', '㎏/3.3㎡'),
        ('잔존 CO₂(ppm)', '잔존CO2', 'CO₂', 'ppm'),
        ('급액 EC(dS/m)', '급액EC', '급액EC', 'dS/m'),
        ('급액 pH', '급액pH', '급액pH', ''),
        ('1회 급액량(㏄/회)', '1회급액량', '1회급액량', '㏄/회'),
        ('1일 공급량(㏄/day)', '1일공급량', '1일공급량', '㏄/day')
    ]
    
    # 유효한 비교 데이터 수집
    valid_comparisons = []
    
    for ref_col, user_key, display_name, unit in all_comparisons:
        user_col = column_mapping.get(user_key)
        
        # 적정생육표 범위 계산
        ref_min, ref_max = None, None
        if ref_col in reference_df.columns:
            try:
                ref_data = pd.to_numeric(reference_df[ref_col], errors='coerce').dropna()
                if len(ref_data) > 0:
                    ref_min = ref_data.min()
                    ref_max = ref_data.max()
            except Exception:
                pass
        
        # 사용자 데이터 평균 계산
        user_avg = None
        if user_col and user_col in user_df.columns:
            try:
                user_data = pd.to_numeric(user_df[user_col], errors='coerce').dropna()
                if len(user_data) > 0:
                    user_avg = user_data.mean()
            except Exception:
                pass
        
        # 둘 다 있을 때만 비교에 포함
        if ref_min is not None and ref_max is not None and user_avg is not None:
            valid_comparisons.append({
                'name': display_name,
                'unit': unit,
                'ref_min': ref_min,
                'ref_max': ref_max,
                'user_avg': user_avg
            })
    
    # 게이지 차트들 생성
    gauge_figs = []
    
    for comparison in valid_comparisons:
        name = comparison['name']
        unit = comparison['unit']
        ref_min = comparison['ref_min']
        ref_max = comparison['ref_max']
        user_avg = comparison['user_avg']
        
        # 게이지 차트 범위 설정 (적정 범위보다 약간 넓게)
        range_margin = (ref_max - ref_min) * 0.2
        gauge_min = max(0, ref_min - range_margin)
        gauge_max = ref_max + range_margin
        
        # 사용자 값이 적정 범위 안에 있는지 확인
        if ref_min <= user_avg <= ref_max:
            gauge_color = "green"
            status = "적정"
        elif user_avg < ref_min:
            gauge_color = "orange"
            status = "낮음"
        else:
            gauge_color = "red"
            status = "높음"
        
        fig = go.Figure(go.Indicator(
            mode = "gauge+number",
            value = user_avg,
            domain = {'x': [0, 1], 'y': [0, 1]},
            title = {'text': f"{name}<br><span style='font-size:0.8em;color:gray'>{status}</span>"},
            gauge = {
                'axis': {'range': [None, gauge_max]},
                'bar': {'color': gauge_color},
                'steps': [
                    {'range': [gauge_min, ref_min], 'color': "lightgray"},
                    {'range': [ref_min, ref_max], 'color': "lightgreen"},
                    {'range': [ref_max, gauge_max], 'color': "lightgray"}
                ],
                'threshold': {
                    'line': {'color': "red", 'width': 4},
                    'thickness': 0.75,
                    'value': user_avg
                }
            },
            number = {'suffix': unit}
        ))
        
        fig.update_layout(
            height=300,
            margin=dict(t=80, b=20, l=20, r=20),
            font={'size': 12}
        )
        
        # 범위 정보 추가
        fig.add_annotation(
            text=f"적정범위: {ref_min:.1f} ~ {ref_max:.1f}{unit}",
            xref="paper", yref="paper",
            x=0.5, y=-0.1,
            showarrow=False,
            font=dict(size=10, color="gray")
        )
        
        gauge_figs.append(fig)
    
    return gauge_figs, valid_comparisons

# 사이드바 - 3단계 선택
st.sidebar.header("📋 선택 옵션")

# 1단계: 일사량/생육상태별 선택
step1 = st.sidebar.selectbox(
    "1단계: 분류 선택",
    ["일사량별", "생육상태별"]
)

# 2단계: 비닐/유리 선택
step2 = st.sidebar.selectbox(
    "2단계: 온실 유형 선택",
    ["비닐", "유리"]
)

# 선택된 Excel 파일 경로
selected_file = EXCEL_FILES[step1][step2]
file_path = os.path.join(DATA_DIR, selected_file)

# 파일 존재 확인
if os.path.exists(file_path):
    # 3단계: 시트명 선택
    sheet_names = get_sheet_names(file_path)
    if sheet_names:
        step3 = st.sidebar.selectbox(
            "3단계: 시트 선택",
            sheet_names
        )
    else:
        st.error("시트를 읽을 수 없습니다.")
        step3 = None
else:
    st.error(f"파일을 찾을 수 없습니다: {file_path}")
    step3 = None

st.sidebar.markdown("---")

# 메인 화면 - 사용자 데이터 업로드만 표시
if step3:
    try:
        # 선택된 시트 로드 (백그라운드에서 처리)
        reference_df = pd.read_excel(file_path, sheet_name=step3)
    except Exception as e:
        st.error(f"데이터 로드 오류: {e}")
        reference_df = None
else:
    reference_df = None

st.header("📁 사용자 데이터 업로드")

# CSV/Excel 파일 업로드
uploaded_file = st.file_uploader(
    "CSV 또는 Excel 파일을 업로드하세요",
    type=['csv', 'xlsx', 'xls'],
    help="누적일사량, 외기기온, 온도, 습도 데이터가 포함된 CSV 또는 Excel 파일"
)

if uploaded_file is not None:
    try:
        # 파일 확장자에 따라 읽기 방법 결정
        file_extension = uploaded_file.name.split('.')[-1].lower()
        
        if file_extension == 'csv':
            user_df = pd.read_csv(uploaded_file)
        elif file_extension in ['xlsx', 'xls']:
            # Excel 파일의 경우 시트 선택 옵션 제공
            excel_file = pd.ExcelFile(uploaded_file)
            sheet_names = excel_file.sheet_names
            
            if len(sheet_names) > 1:
                selected_sheet = st.selectbox(
                    "Excel 시트를 선택하세요:",
                    sheet_names,
                    help="분석할 데이터가 있는 시트를 선택하세요"
                )
            else:
                selected_sheet = sheet_names[0]
            
            user_df = pd.read_excel(uploaded_file, sheet_name=selected_sheet)
        else:
            st.error("지원하지 않는 파일 형식입니다.")
            user_df = None
        
        # 필수 컬럼 검증 (user_df가 None이 아닌 경우에만)
        if user_df is not None:
            is_valid, missing_columns, column_mapping = validate_user_data(user_df)
        else:
            is_valid = False
            missing_columns = []
            column_mapping = {}
            
        if user_df is not None and is_valid:
            column_info = ", ".join([f"{col_name}" for col_name in column_mapping])
            st.success(f"발견된 분석 가능 컬럼: {column_info}")
                            
            # 매칭 실행 버튼
            if st.button("🔍 구간 매칭 및 비교 분석", type="primary"):
                if reference_df is not None:
                    with st.spinner("데이터 분석 중..."):
                        try:
                            # 구간 매칭
                            matched_reference, user_data = find_matching_range(reference_df, user_df, column_mapping)
                            
                            if not matched_reference.empty:
                                st.success(f"✅ 매칭 완료! 적정생육표에서 {len(matched_reference)}개의 해당 구간을 찾았습니다.")
                                
                                # 매칭된 구간 결과 표시
                                st.subheader("📊 매칭된 적정생육표 구간")
                                st.dataframe(matched_reference.reset_index(drop=True), use_container_width=True)
                                
                                # 게이지 차트로 데이터 비교 분석
                                st.subheader("📈 데이터 비교 분석")
                                
                                gauge_figs, valid_comparisons = create_comparison_charts(matched_reference, user_data, column_mapping)
                                
                                if gauge_figs:
                                    # 게이지 차트들을 2개씩 한 줄로 배치
                                    for i in range(0, len(gauge_figs), 2):
                                        cols = st.columns(2)
                                        
                                        # 첫 번째 게이지
                                        with cols[0]:
                                            st.plotly_chart(gauge_figs[i], use_container_width=True)
                                        
                                        # 두 번째 게이지 (있는 경우)
                                        if i + 1 < len(gauge_figs):
                                            with cols[1]:
                                                st.plotly_chart(gauge_figs[i + 1], use_container_width=True)
                                    
                                    # 요약 정보
                                    st.subheader("📋 분석 요약")
                                    
                                    # 상태별 개수 집계
                                    status_counts = {'적정': 0, '낮음': 0, '높음': 0}
                                    for comparison in valid_comparisons:
                                        ref_min = comparison['ref_min']
                                        ref_max = comparison['ref_max']
                                        user_avg = comparison['user_avg']
                                        
                                        if ref_min <= user_avg <= ref_max:
                                            status_counts['적정'] += 1
                                        elif user_avg < ref_min:
                                            status_counts['낮음'] += 1
                                        else:
                                            status_counts['높음'] += 1
                                    
                                    # 상태 요약 표시
                                    col1, col2, col3 = st.columns(3)
                                    with col1:
                                        st.metric("🟢 적정 범위", f"{status_counts['적정']}개")
                                    with col2:
                                        st.metric("🟡 기준 미달", f"{status_counts['낮음']}개")
                                    with col3:
                                        st.metric("🔴 기준 초과", f"{status_counts['높음']}개")
                                    
                                    # 전체 적정성 판단
                                    total_items = len(valid_comparisons)
                                    optimal_ratio = status_counts['적정'] / total_items * 100
                                    
                                    if optimal_ratio >= 80:
                                        st.success(f"✅ 전반적으로 양호합니다! ({optimal_ratio:.0f}% 적정)")
                                    elif optimal_ratio >= 60:
                                        st.warning(f"⚠️ 일부 개선이 필요합니다. ({optimal_ratio:.0f}% 적정)")
                                    else:
                                        st.error(f"🚨 전반적인 개선이 필요합니다. ({optimal_ratio:.0f}% 적정)")
                                
                                else:
                                    st.info("비교 가능한 데이터가 없습니다. 사용자 데이터에 적정생육표와 매칭되는 컬럼이 있는지 확인해주세요.")
                                
                                    
                            else:
                                st.warning("❌ 사용자 데이터에 해당하는 적정생육표 구간이 없습니다.")
                                st.info("다른 시트나 분류를 선택해보세요.")
                        except Exception as e:
                            st.error(f"분석 오류: {e}")
                else:
                    st.error("적정생육표를 먼저 선택해주세요.")
        elif user_df is not None:
            st.error("❌ 분석 가능한 컬럼이 없습니다.")
            st.info("현재 파일의 컬럼: " + ", ".join(user_df.columns.tolist()))
                    
    except Exception as e:
        st.error(f"파일 읽기 오류: {e}")

# 사용법 안내
st.markdown("---")
st.header("💡 사용법")
st.markdown("""
1. **사이드바에서 3단계 선택**: 일사량/생육상태별 → 비닐/유리 → 시트명
2. **데이터 업로드**: CSV/Excel 파일 (누적일사량, 외기기온, 온도, 습도 등 포함)
3. **분석 실행**: '구간 매칭 및 비교 분석' 버튼 클릭
4. **결과 확인**: 매칭된 구간, 비교 그래프
""")
