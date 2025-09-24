import streamlit as st
import pandas as pd
import os
import plotly.graph_objects as go

# 페이지 설정
st.set_page_config(
    page_title="토마토 최적환경 설정구간 매칭시스템",
    page_icon="🍅",
    layout="centered",  # 모바일 친화적으로 변경
    initial_sidebar_state="collapsed"  # 모바일에서 사이드바 접힘 상태로 시작
)


# 모바일 친화적 CSS 스타일 추가
st.markdown("""
<style>
    /* 모바일 반응형 스타일 */
    @media (max-width: 768px) {
        .main .block-container {
            padding-top: 1rem;
            padding-left: 1rem;
            padding-right: 1rem;
        }
        
        .stSelectbox > div > div {
            font-size: 14px;
        }
        
        .stButton > button {
            width: 100%;
            margin-top: 10px;
        }
        
        /* 테이블 스크롤 가능하게 */
        .dataframe {
            overflow-x: auto;
        }
    }
    
    /* 게이지 차트 모바일 최적화 */
    .js-plotly-plot {
        width: 100% !important;
    }
    
    /* 게이지 차트 제목 최적화 */
    .js-plotly-plot .gtitle {
        font-size: 14px !important;
        line-height: 1.2 !important;
    }
    
    /* Plotly 차트 컨테이너 최적화 */
    .plotly-graph-div {
        overflow: visible !important;
    }
</style>
""", unsafe_allow_html=True)


st.title("🍅 토마토 최적환경 설정구간 매칭시스템")
st.markdown("---")

# 데이터 파일 경로
DATA_DIR = "data"

# Excel 파일 목록 (기본적으로 일사량별 사용)
EXCEL_FILES = {
    "비닐": "solar_level_vinyl_standard.xlsx",
    "유리": "solar_level_glass_standard.xlsx"
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
    # 실제 Excel 파일의 정확한 컬럼명 매핑 (여러 가능한 형태 포함)
    exact_columns_map = {
        '누적일사량': ['누적일사량(범위)', '누적일사량(MJ/m²기간평균)', '누적일사량'],
        '외기기온': ['외기기온(범위)', '외기기온(℃)', '외기기온'],
        '생산량': ['생산량(㎏/3.3㎡)', '생산량(kg/3.3㎡)', '생산량'],
        '일일평균온도': ['일일 평균온도(℃)', '일일평균온도(℃)', '일일평균온도'],
        '주간평균온도': ['주간 평균온도(℃)', '주간평균온도(℃)', '주간평균온도'],
        '야간평균온도': ['야간 평균온도(℃)', '야간평균온도(℃)', '야간평균온도'],
        '새벽온도': ['새벽온도(℃)', '새벽온도'],
        '주간평균습도': ['주간 평균습도(%)', '주간평균습도(%)', '주간평균습도'],
        '잔존CO2': ['잔존 CO₂(ppm)', '잔존CO2(ppm)', '잔존CO2'],
        '급액EC': ['급액 EC(dS/m)', '급액EC(dS/m)', '급액EC'],
        '급액pH': ['급액 pH', '급액pH'],
        '1회급액량': ['1회 급액량(㏄/회)', '1회급액량(cc/회)', '1회급액량'],
        '1일공급량': ['1일 공급량(㏄/day)', '1일공급량(cc/day)', '1일공급량']
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
    반환: (matched_reference, user_data, matching_info)
    """
    # 컬럼 매핑에서 실제 컬럼명 가져오기
    radiation_col = column_mapping.get('누적일사량')
    temp_col = column_mapping.get('외기기온')
    
    # 적정생육표의 해당 컬럼 확인 (정확한 컬럼명)
    ref_radiation_col = None
    ref_temp_col = None
    
    # 적정생육표에서 누적일사량 컬럼 찾기 (다양한 형태 지원)
    radiation_possible_names = ['누적일사량(범위)', '누적일사량(MJ/m²기간평균)', '누적일사량']
    for possible_name in radiation_possible_names:
        if possible_name in reference_df.columns:
            ref_radiation_col = possible_name
            break
    
    # 적정생육표에서 외기기온 컬럼 찾기 (다양한 형태 지원)
    temp_possible_names = ['외기기온(범위)', '외기기온(℃)', '외기기온']
    for possible_name in temp_possible_names:
        if possible_name in reference_df.columns:
            ref_temp_col = possible_name
            break
    
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
                # 사용자 누적일사량 정보를 변수에 저장 (표시하지 않음)
                if user_radiation_min == user_radiation_max:
                    user_radiation_info = f"{user_radiation_min:.2f}"
                else:
                    user_radiation_info = f"{user_radiation_min:.2f} ~ {user_radiation_max:.2f}"
        except Exception as e:
            st.warning(f"누적일사량 데이터 처리 중 오류: {e}")
    
    # 사용자 데이터에서 외기기온 범위 계산
    if temp_col and temp_col in user_data.columns:
        try:
            user_temp_data = pd.to_numeric(user_data[temp_col], errors='coerce').dropna()
            if len(user_temp_data) > 0:
                user_temp_min = user_temp_data.min()
                user_temp_max = user_temp_data.max()
                # 사용자 외기기온 정보를 변수에 저장 (표시하지 않음)
                if user_temp_min == user_temp_max:
                    user_temp_info = f"{user_temp_min:.2f}°C"
                else:
                    user_temp_info = f"{user_temp_min:.2f} ~ {user_temp_max:.2f}°C"
        except Exception as e:
            st.warning(f"외기기온 데이터 처리 중 오류: {e}")
    
    # 매칭 정보 수집을 위한 변수들
    matching_info = {
        'user_radiation_range': (user_radiation_min, user_radiation_max) if user_radiation_min is not None else None,
        'user_temp_range': (user_temp_min, user_temp_max) if user_temp_min is not None else None,
        'user_radiation_info': user_radiation_info if 'user_radiation_info' in locals() else None,
        'user_temp_info': user_temp_info if 'user_temp_info' in locals() else None,
        'available_radiation_ranges': [],
        'available_temp_ranges': [],
        'radiation_matches': [],
        'temp_matches': [],
        'both_conditions_checked': bool(ref_radiation_col and ref_temp_col and 
                                      user_radiation_min is not None and user_temp_min is not None)
    }
    
    # 적정생육표에서 사용자 범위를 포함하는 구간 찾기
    matched_rows = []
    
    for idx, row in reference_df.iterrows():
        is_match = True
        
        # 누적일사량 범위 체크
        radiation_match = True
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
                        
                        # 적정생육표 범위 수집
                        matching_info['available_radiation_ranges'].append((ref_min, ref_max))
                        
                        # 적정생육표 범위가 사용자 범위를 포함하는지 확인
                        if ref_min <= user_radiation_min and ref_max >= user_radiation_max:
                            matching_info['radiation_matches'].append((ref_min, ref_max))
                        else:
                            radiation_match = False
                else:
                    # 단일 값인 경우
                    ref_val = float(ref_radiation_str)
                    matching_info['available_radiation_ranges'].append((ref_val, ref_val))
                    
                    if user_radiation_min <= ref_val <= user_radiation_max:
                        matching_info['radiation_matches'].append((ref_val, ref_val))
                    else:
                        radiation_match = False
            except (ValueError, TypeError):
                # 파싱 실패 시 해당 행은 제외
                radiation_match = False
        
        if not radiation_match:
                is_match = False
        
        # 외기기온 범위 체크
        temp_match = True
        if ref_temp_col and user_temp_min is not None and user_temp_max is not None:
            try:
                ref_temp_str = str(row[ref_temp_col])
                # 범위 형태 파싱
                if '-' in ref_temp_str or '~' in ref_temp_str:
                    separator = '-' if '-' in ref_temp_str else '~'
                    parts = ref_temp_str.split(separator)
                    if len(parts) == 2:
                        ref_min = float(parts[0].strip())
                        ref_max = float(parts[1].strip())
                        
                        # 적정생육표 범위 수집
                        matching_info['available_temp_ranges'].append((ref_min, ref_max))
                        
                        # 적정생육표 범위가 사용자 범위를 포함하는지 확인
                        if ref_min <= user_temp_min and ref_max >= user_temp_max:
                            matching_info['temp_matches'].append((ref_min, ref_max))
                        else:
                            temp_match = False
                else:
                    # 단일 값인 경우
                    ref_val = float(ref_temp_str)
                    matching_info['available_temp_ranges'].append((ref_val, ref_val))
                    
                    if user_temp_min <= ref_val <= user_temp_max:
                        matching_info['temp_matches'].append((ref_val, ref_val))
                    else:
                        temp_match = False
            except (ValueError, TypeError):
                temp_match = False
        
        if not temp_match:
                is_match = False
        
        if is_match:
            matched_rows.append(row)
    
    if matched_rows:
        matched_reference = pd.DataFrame(matched_rows)
    else:
        matched_reference = pd.DataFrame()
    
    # 중복 제거
    matching_info['available_radiation_ranges'] = list(set(matching_info['available_radiation_ranges']))
    matching_info['available_temp_ranges'] = list(set(matching_info['available_temp_ranges']))
    matching_info['radiation_matches'] = list(set(matching_info['radiation_matches']))
    matching_info['temp_matches'] = list(set(matching_info['temp_matches']))
    
    return matched_reference, user_data, matching_info

def create_comparison_charts(reference_df: pd.DataFrame, user_df: pd.DataFrame, column_mapping: dict, max_production_row=None):
    """
    게이지 차트로 사용자 평균값과 적정생육표 범위 비교
    환경 데이터와 생산량 데이터를 분리하여 반환
    """
    import plotly.graph_objects as go
    
    def find_column_in_df(df, possible_names):
        """데이터프레임에서 가능한 컬럼명 중 실제 존재하는 컬럼 찾기"""
        for name in possible_names:
            if name in df.columns:
                return name
        return None
    
    # 환경 데이터 비교 항목들 (실제 존재하는 컬럼만 사용)
    environment_comparisons = []
    env_column_mapping = {
        '일일평균온도': ['일일 평균온도(℃)', '일일평균온도(℃)', '일일평균온도'],
        '주간평균온도': ['주간 평균온도(℃)', '주간평균온도(℃)', '주간평균온도'],
        '야간평균온도': ['야간 평균온도(℃)', '야간평균온도(℃)', '야간평균온도'],
        '새벽온도': ['새벽온도(℃)', '새벽온도'],
        '주간평균습도': ['주간 평균습도(%)', '주간평균습도(%)', '주간평균습도'],
        '잔존CO2': ['잔존 CO₂(ppm)', '잔존CO2(ppm)', '잔존CO2'],
        '급액EC': ['급액 EC(dS/m)', '급액EC(dS/m)', '급액EC'],
        '급액pH': ['급액 pH', '급액pH'],
        '1회급액량': ['1회 급액량(㏄/회)', '1회급액량(cc/회)', '1회급액량'],
        '1일공급량': ['1일 공급량(㏄/day)', '1일공급량(cc/day)', '1일공급량']
    }
    
    # 실제 존재하는 환경 데이터 컬럼만 추가
    for key, possible_names in env_column_mapping.items():
        found_col = find_column_in_df(reference_df, possible_names)
        if found_col:
            unit_map = {
                '일일평균온도': '°C', '주간평균온도': '°C', '야간평균온도': '°C', '새벽온도': '°C',
                '주간평균습도': '%', '잔존CO2': 'ppm', '급액EC': 'dS/m', '급액pH': '',
                '1회급액량': '㏄/회', '1일공급량': '㏄/day'
            }
            environment_comparisons.append((found_col, key, key, unit_map.get(key, '')))
    
    # 생산량 데이터 비교 항목들
    production_comparisons = []
    prod_possible_names = ['생산량(㎏/3.3㎡)', '생산량(kg/3.3㎡)', '생산량']
    found_prod_col = find_column_in_df(reference_df, prod_possible_names)
    if found_prod_col:
        production_comparisons.append((found_prod_col, '생산량', '생산량', '㎏/3.3㎡'))
    
    def process_comparisons(comparison_list, category_name):
        """특정 카테고리의 비교 데이터 처리"""
        valid_comparisons = []
        gauge_figs = []
        
        for ref_col, user_key, display_name, unit in comparison_list:
            user_col = column_mapping.get(user_key)
            
            # 적정생육표 범위 계산 (최대 생산량 행 기준)
            ref_min, ref_max = None, None
            if ref_col in reference_df.columns:
                try:
                    if max_production_row is not None and ref_col in max_production_row.index:
                        # 최대 생산량 행의 값을 최소/최대로 사용 (범위 데이터 파싱)
                        ref_value_str = str(max_production_row[ref_col])
                        
                        # 범위 형태 파싱 (예: "20~25", "20-25")
                        if '~' in ref_value_str or '-' in ref_value_str:
                            separator = '~' if '~' in ref_value_str else '-'
                            parts = ref_value_str.split(separator)
                            if len(parts) == 2:
                                try:
                                    ref_min = float(parts[0].strip())
                                    ref_max = float(parts[1].strip())
                                except (ValueError, TypeError):
                                    pass
                        else:
                            # 단일 값인 경우
                            try:
                                ref_value = float(ref_value_str)
                                ref_min = ref_value
                                ref_max = ref_value
                            except (ValueError, TypeError):
                                pass
                    else:
                        # 기존 방식: 전체 데이터의 최소/최대값 사용
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
        for comparison in valid_comparisons:
            name = comparison['name']
            unit = comparison['unit']
            ref_min = comparison['ref_min']
            ref_max = comparison['ref_max']
            user_avg = comparison['user_avg']
            
            # 게이지 차트 범위 설정 (적정 범위보다 약간 넓게)
            if ref_min == ref_max:
                # 최대 생산량 행의 단일 값인 경우
                range_margin = abs(ref_min) * 0.2 if ref_min != 0 else 1
                gauge_min = max(0, ref_min - range_margin)
                gauge_max = ref_max + range_margin
            else:
                # 범위가 있는 경우
                range_margin = (ref_max - ref_min) * 0.2
                gauge_min = max(0, ref_min - range_margin)
                gauge_max = ref_max + range_margin
            
            # 사용자 값이 적정 범위 안에 있는지 확인
            if ref_min == ref_max:
                # 단일 값과 비교 (정확한 비교, 허용 오차 없음)
                if user_avg == ref_min:
                    gauge_color = "green"
                    status = "적정"
                elif user_avg < ref_min:
                    gauge_color = "orange"
                    status = "낮음"
                else:
                    gauge_color = "red"
                    status = "높음"
            else:
                # 범위와 비교 (최대 생산량 행의 환경 조건 범위)
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
                title = {
                    'text': f"{name}<br><span style='font-size:0.9em;color:gray'>{status}</span>",
                    'font': {'size': 14}
                },
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
                height=300,  # 제목 공간 확보를 위해 높이 증가
                margin=dict(t=100, b=40, l=15, r=15),  # 상단 마진 크게 증가하여 제목 공간 확보
                font={'size': 11}  # 폰트 크기 약간 증가
            )
            
            # 범위 정보 추가
            if ref_min == ref_max:
                range_text = f"최적값: {ref_min:.1f}{unit}"
            else:
                range_text = f"적정범위: {ref_min:.1f} ~ {ref_max:.1f}{unit}"
            
            fig.add_annotation(
                text=range_text,
                xref="paper", yref="paper",
                x=0.5, y=-0.15,  # 위치를 더 아래로 이동
                showarrow=False,
                font=dict(size=11, color="gray")
            )
            
            gauge_figs.append(fig)
        
        return gauge_figs, valid_comparisons

    # 환경 데이터와 생산량 데이터 각각 처리
    env_gauge_figs, env_valid_comparisons = process_comparisons(environment_comparisons, "환경")
    prod_gauge_figs, prod_valid_comparisons = process_comparisons(production_comparisons, "생산량")
    
    return env_gauge_figs, env_valid_comparisons, prod_gauge_figs, prod_valid_comparisons

# 사이드바 - 2단계 선택
st.sidebar.header("📋 선택 옵션")

# 1단계: 온실 유형 선택
step1 = st.sidebar.selectbox(
    "1단계: 온실 유형 선택",
    ["비닐", "유리"]
)

# 선택된 Excel 파일 경로
selected_file = EXCEL_FILES[step1]
file_path = os.path.join(DATA_DIR, selected_file)

# 파일 존재 확인
if os.path.exists(file_path):
    # 2단계: 날짜 선택 (실제로는 시트 선택)
    sheet_names = get_sheet_names(file_path)
    if sheet_names:
        step2 = st.sidebar.selectbox(
            "2단계: 날짜 선택",
            sheet_names
        )
        file_exists = True
    else:
        st.error("시트를 읽을 수 없습니다.")
        step2 = None
        file_exists = False
else:
    st.error(f"파일을 찾을 수 없습니다: {file_path}")
    step2 = None
    file_exists = False

st.sidebar.markdown("---")

# 메인 화면 - 사용자 데이터 업로드만 표시
if file_exists and step2:
    try:
        # 선택된 시트 로드
        reference_df = pd.read_excel(file_path, sheet_name=step2)
    except Exception as e:
        st.error(f"데이터 로드 오류: {e}")
        reference_df = None
else:
    reference_df = None

st.header("📂 내 농가 데이터 업로드")

# 예시 파일 다운로드 기능
st.subheader("📥 예시 파일 다운로드")

# 예시 파일 경로
example_file_path = "data/test_tomato_data.xlsx"

try:
    if os.path.exists(example_file_path):
        with open(example_file_path, "rb") as file:
            example_file_data = file.read()
        
        st.download_button(
            label="📊 예시 파일 다운로드 (test_tomato_data.xlsx)",
            data=example_file_data,
            file_name="test_tomato_data.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            help="토마토 환경 데이터 예시 파일을 다운로드하여 참고하세요"
        )
        st.info("💡 위 예시 파일을 다운로드하여 데이터 형식을 참고하거나 테스트용으로 사용하실 수 있습니다.")
    else:
        st.warning("예시 파일을 찾을 수 없습니다.")
except Exception as e:
    st.error(f"예시 파일 로드 중 오류: {e}")

st.markdown("---")

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
                            matched_reference, user_data, matching_info = find_matching_range(reference_df, user_df, column_mapping)
                            
                            if not matched_reference.empty:
                                st.success(f"✅ 매칭 완료! 적정생육표에서 {len(matched_reference)}개의 해당 구간을 찾았습니다.")
                                
                                # 매칭된 구간 결과 표시
                                st.subheader("📊 최적 환경표")
                                
                                # 사용자 데이터 정보를 한 줄로 표시
                                user_radiation_display = matching_info.get('user_radiation_info', '정보 없음')
                                user_temp_display = matching_info.get('user_temp_info', '정보 없음')
                                if user_radiation_display != '정보 없음' and user_temp_display != '정보 없음':
                                    st.info(f"📊 사용자 누적일사량: {user_radiation_display} | 🌡️ 사용자 외기기온: {user_temp_display}")
                                elif user_radiation_display != '정보 없음':
                                    st.info(f"📊 사용자 누적일사량: {user_radiation_display}")
                                elif user_temp_display != '정보 없음':
                                    st.info(f"🌡️ 사용자 외기기온: {user_temp_display}")
                                
                                # 생산량이 가장 높은 행 찾기
                                production_col = '생산량(㎏/3.3㎡)'
                                max_production_idx = None
                                max_production_value = None
                                
                                if production_col in matched_reference.columns:
                                    try:
                                        # 생산량 컬럼을 숫자로 변환
                                        production_data = pd.to_numeric(matched_reference[production_col], errors='coerce')
                                        if not production_data.isna().all():
                                            max_production_idx = production_data.idxmax()
                                            max_production_value = production_data.max()
                                    except Exception as e:
                                        st.warning(f"생산량 데이터 처리 중 오류: {e}")
                                
                                # 스타일이 적용된 데이터프레임 표시
                                if max_production_idx is not None:
                                    # 인덱스를 리셋한 후 최대 생산량 행의 새로운 인덱스 찾기
                                    reset_df = matched_reference.reset_index(drop=True)
                                    
                                    # 원본 데이터프레임에서 최대 생산량 행의 위치 찾기
                                    max_row_in_original = matched_reference.loc[max_production_idx]
                                    
                                    # 리셋된 데이터프레임에서 같은 행 찾기
                                    new_max_idx = None
                                    for idx, row in reset_df.iterrows():
                                        if row.equals(max_row_in_original):
                                            new_max_idx = idx
                                            break
                                    
                                    # 스타일링 함수
                                    def highlight_max_production(df):
                                        """생산량이 최대인 행을 노란색으로 강조"""
                                        style_df = pd.DataFrame('', index=df.index, columns=df.columns)
                                        if new_max_idx is not None:
                                            style_df.loc[new_max_idx] = 'background-color: yellow'
                                        return style_df
                                    
                                    styled_df = reset_df.style.apply(highlight_max_production, axis=None)
                                    st.dataframe(styled_df, use_container_width=True)
                                else:
                                    st.dataframe(matched_reference.reset_index(drop=True), use_container_width=True)
                                
                                # 최대 생산량 정보 표시
                                if max_production_idx is not None and max_production_value is not None:
                                    max_row = matched_reference.loc[max_production_idx]
                                    
                                    # 누적일사량과 외기기온 정보 추출 및 정리
                                    radiation_info = "정보 없음"
                                    temp_info = "정보 없음"
                                    
                                    if '누적일사량(범위)' in max_row.index:
                                        radiation_raw = str(max_row['누적일사량(범위)'])
                                        # 숫자와 범위 구분자만 추출하여 정리
                                        if '~' in radiation_raw or '-' in radiation_raw:
                                            # 범위 형태인 경우
                                            separator = '~' if '~' in radiation_raw else '-'
                                            parts = radiation_raw.split(separator)
                                            if len(parts) == 2:
                                                try:
                                                    min_val = float(parts[0].strip())
                                                    max_val = float(parts[1].strip())
                                                    radiation_info = f"{min_val:.0f}~{max_val:.0f}"
                                                except:
                                                    radiation_info = radiation_raw
                                            else:
                                                radiation_info = radiation_raw
                                        else:
                                            # 단일 값인 경우
                                            try:
                                                val = float(radiation_raw)
                                                radiation_info = f"{val:.0f}"
                                            except:
                                                radiation_info = radiation_raw
                                    
                                    if '외기기온(범위)' in max_row.index:
                                        temp_raw = str(max_row['외기기온(범위)'])
                                        # 숫자와 범위 구분자만 추출하여 정리
                                        if '~' in temp_raw or '-' in temp_raw:
                                            # 범위 형태인 경우
                                            separator = '~' if '~' in temp_raw else '-'
                                            parts = temp_raw.split(separator)
                                            if len(parts) == 2:
                                                try:
                                                    min_val = float(parts[0].strip())
                                                    max_val = float(parts[1].strip())
                                                    temp_info = f"{min_val:.1f}~{max_val:.1f}°C"
                                                except:
                                                    temp_info = temp_raw + "°C"
                                            else:
                                                temp_info = temp_raw + "°C"
                                        else:
                                            # 단일 값인 경우
                                            try:
                                                val = float(temp_raw)
                                                temp_info = f"{val:.1f}°C"
                                            except:
                                                temp_info = temp_raw + "°C"
                                    
                                    st.info(f"🏆 사용자의 누적일사량, 외기기온 구간의 조건에서 최대 **{max_production_value:.1f}㎏/3.3㎡** 수확 가능")
                                
                                # 환경 데이터 비교 분석 (최대 생산량 행 기준)
                                max_production_row_data = None
                                if max_production_idx is not None:
                                    max_production_row_data = matched_reference.loc[max_production_idx]
                                env_gauge_figs, env_valid_comparisons, _, _ = create_comparison_charts(matched_reference, user_data, column_mapping, max_production_row_data)
                                
                                # 최고 수확량 기준 내 농가 비교 분석을 먼저 표시
                                if env_valid_comparisons:
                                    st.subheader("📊 내 농가 환경 진단")
                                    
                                    # 환경 데이터 상태별 개수 집계
                                    env_status_counts = {'적정': 0, '낮음': 0, '높음': 0}
                                    for comparison in env_valid_comparisons:
                                        ref_min = comparison['ref_min']
                                        ref_max = comparison['ref_max']
                                        user_avg = comparison['user_avg']
                                        
                                        if ref_min == ref_max:
                                            # 단일 값과 비교 (정확한 비교, 허용 오차 없음)
                                            if user_avg == ref_min:
                                                env_status_counts['적정'] += 1
                                            elif user_avg < ref_min:
                                                env_status_counts['낮음'] += 1
                                            else:
                                                env_status_counts['높음'] += 1
                                        else:
                                            # 범위와 비교
                                            if ref_min <= user_avg <= ref_max:
                                                env_status_counts['적정'] += 1
                                            elif user_avg < ref_min:
                                                env_status_counts['낮음'] += 1
                                            else:
                                                env_status_counts['높음'] += 1
                                    
                                    # 환경 상태 요약 표시
                                    col1, col2, col3 = st.columns(3)
                                    with col1:
                                        st.metric("🟢 적정 범위", f"{env_status_counts['적정']}개")
                                    with col2:
                                        st.metric("🟡 기준 미달", f"{env_status_counts['낮음']}개")
                                    with col3:
                                        st.metric("🔴 기준 초과", f"{env_status_counts['높음']}개")
                                    
                                    # 각 변수별 상태 상세 표시
                                    if len(env_valid_comparisons) > 0:  # 비교 데이터가 있는 경우에만 표시
                                        st.subheader("🔍 환경 요소별 상세 분석")
                                        
                                        low_items = []
                                        high_items = []
                                        optimal_items = []
                                        
                                        for comparison in env_valid_comparisons:
                                            name = comparison['name']
                                            ref_min = comparison['ref_min']
                                            ref_max = comparison['ref_max']
                                            user_avg = comparison['user_avg']
                                            
                                            if ref_min == ref_max:
                                                # 최대 생산량 행의 단일 값과 비교 (정확한 비교, 허용 오차 없음)
                                                if user_avg == ref_min:
                                                    optimal_items.append(name)
                                                elif user_avg < ref_min:
                                                    gap = ref_min - user_avg
                                                    low_items.append(f"**{name}**: {gap:.1f} 부족")
                                                else:
                                                    gap = user_avg - ref_min
                                                    high_items.append(f"**{name}**: {gap:.1f} 초과")
                                            else:
                                                # 범위와 비교
                                                if ref_min <= user_avg <= ref_max:
                                                    optimal_items.append(name)
                                                elif user_avg < ref_min:
                                                    gap = ref_min - user_avg
                                                    low_items.append(f"**{name}**: {gap:.1f} 부족")
                                                else:
                                                    gap = user_avg - ref_max
                                                    high_items.append(f"**{name}**: {gap:.1f} 초과")
                                        
                                        # 3개 컬럼으로 분류하여 표시 (적정 범위 → 기준 미달 → 기준 초과 순서)
                                        col1, col2, col3 = st.columns(3)
                                        
                                        with col1:
                                            if optimal_items:
                                                st.markdown("🟢 **적정 범위**")
                                                for item in optimal_items:
                                                    st.markdown(f"• **{item}**: 양호")
                                            
                                        with col2:
                                            if low_items:
                                                st.markdown("🟡 **기준 미달**")
                                                for item in low_items:
                                                    st.markdown(f"• {item}")
                                        
                                        with col3:
                                            if high_items:
                                                st.markdown("🔴 **기준 초과**")
                                                for item in high_items:
                                                    st.markdown(f"• {item}")

                                # 환경 데이터 비교 분석
                                if env_gauge_figs:
                                    st.subheader("📈 환경 데이터 게이지 차트")
                                    
                                    # 모바일 화면 크기 감지를 위한 JavaScript 대신 간단한 방법 사용
                                    use_single_column = len(env_gauge_figs) <= 3 or st.sidebar.button("📱 모바일 모드", help="차트를 세로로 배치합니다")
                                    
                                    if use_single_column:
                                        # 모바일: 1개씩 배치
                                        for fig in env_gauge_figs:
                                            st.plotly_chart(fig, use_container_width=True)
                                    else:
                                        # 데스크톱: 2개씩 배치
                                        for i in range(0, len(env_gauge_figs), 2):
                                            cols = st.columns(2)
                                            
                                            # 첫 번째 게이지
                                            with cols[0]:
                                                st.plotly_chart(env_gauge_figs[i], use_container_width=True)
                                            
                                            # 두 번째 게이지 (있는 경우)
                                            if i + 1 < len(env_gauge_figs):
                                                with cols[1]:
                                                    st.plotly_chart(env_gauge_figs[i + 1], use_container_width=True)
                                
                                
                                
                                if not env_gauge_figs:
                                    st.info("비교 가능한 환경 데이터가 없습니다. 사용자 데이터에 적정생육표와 매칭되는 환경 컬럼이 있는지 확인해주세요.")
                                
                                    
                            else:
                                st.warning("❌ 사용자 데이터에 해당하는 적정생육표 구간이 없습니다.")
                                
                                # 상세한 매칭 실패 정보 표시
                                st.subheader("❌ 매칭 실패 원인 분석")
                                
                                # 사용자 데이터 범위 표시
                                col1, col2 = st.columns(2)
                                
                                with col1:
                                    if matching_info['user_radiation_range']:
                                        user_rad_min, user_rad_max = matching_info['user_radiation_range']
                                        if user_rad_min == user_rad_max:
                                            st.metric("📊 사용자 누적일사량", f"{user_rad_min:.1f}")
                                        else:
                                            st.metric("📊 사용자 누적일사량", f"{user_rad_min:.1f} ~ {user_rad_max:.1f}")
                                        
                                        if matching_info['available_radiation_ranges']:
                                            # 모든 범위에서 최소값과 최대값 추출
                                            all_rad_values = []
                                            for rad_min, rad_max in matching_info['available_radiation_ranges']:
                                                all_rad_values.extend([rad_min, rad_max])
                                            
                                            if all_rad_values:
                                                overall_min = min(all_rad_values)
                                                overall_max = max(all_rad_values)
                                                st.write("**🎯 적정생육표 누적일사량 전체 범위:**")
                                                st.write(f"**{overall_min:.1f} ~ {overall_max:.1f}**")
                                                
                                                # 사용자 값과 비교
                                                if user_rad_max < overall_min:
                                                    gap = overall_min - user_rad_max
                                                    st.error(f"💡 사용자 누적일사량이 적정 범위보다 {gap:.1f} 낮습니다.")
                                                elif user_rad_min > overall_max:
                                                    gap = user_rad_min - overall_max
                                                    st.error(f"💡 사용자 누적일사량이 적정 범위보다 {gap:.1f} 높습니다.")
                                                else:
                                                    st.info("💡 사용자 누적일사량이 적정 범위와 일부 겹칩니다.")
                                        
                                        if not matching_info['radiation_matches']:
                                            st.error("❌ 누적일사량이 적정 범위에 포함되지 않음")
                                        else:
                                            st.success("✅ 누적일사량 매칭됨")
                                
                                with col2:
                                    if matching_info['user_temp_range']:
                                        user_temp_min, user_temp_max = matching_info['user_temp_range']
                                        st.metric("🌡️ 사용자 외기기온", f"{user_temp_min:.1f} ~ {user_temp_max:.1f}°C")
                                        
                                        if matching_info['available_temp_ranges']:
                                            # 모든 범위에서 최소값과 최대값 추출
                                            all_temp_values = []
                                            for temp_min, temp_max in matching_info['available_temp_ranges']:
                                                all_temp_values.extend([temp_min, temp_max])
                                            
                                            if all_temp_values:
                                                overall_min = min(all_temp_values)
                                                overall_max = max(all_temp_values)
                                                st.write("**🎯 적정생육표 외기기온 전체 범위:**")
                                                st.write(f"**{overall_min:.1f} ~ {overall_max:.1f}°C**")
                                                
                                                # 사용자 값과 비교
                                                if user_temp_max < overall_min:
                                                    gap = overall_min - user_temp_max
                                                    st.error(f"💡 사용자 외기기온이 적정 범위보다 {gap:.1f}°C 낮습니다.")
                                                elif user_temp_min > overall_max:
                                                    gap = user_temp_min - overall_max
                                                    st.error(f"💡 사용자 외기기온이 적정 범위보다 {gap:.1f}°C 높습니다.")
                                                else:
                                                    st.info("💡 사용자 외기기온이 적정 범위와 일부 겹칩니다.")
                                        
                                        if not matching_info['temp_matches']:
                                            st.error("❌ 외기기온이 적정 범위에 포함되지 않음")
                                        else:
                                            st.success("✅ 외기기온 매칭됨")
                                
                                # 개선 제안
                                st.subheader("💡 매칭을 위한 개선 방안")
                                if matching_info['both_conditions_checked']:
                                    if not matching_info['radiation_matches'] and not matching_info['temp_matches']:
                                        st.info("🔄 누적일사량과 외기기온 모두 적정 범위를 벗어났습니다. 다른 날짜나 분류를 선택해보세요.")
                                    elif not matching_info['radiation_matches']:
                                        st.info("☀️ 누적일사량이 범위를 벗어났습니다. 일사량별 다른 분류를 확인해보세요.")
                                    elif not matching_info['temp_matches']:
                                        st.info("🌡️ 외기기온이 범위를 벗어났습니다. 생육상태별 다른 분류를 확인해보세요.")
                                else:
                                    st.info("📋 일부 조건만 확인되었습니다. 데이터에 누적일사량과 외기기온 정보가 모두 포함되어 있는지 확인해주세요.")
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
st.header("📖 이용 가이드")

st.markdown("""
**📋 간단 사용법**
1. **기본 설정**: 사이드바에서 온실유형(비닐/유리) 선택 후 분석할 날짜(시트) 선택
2. **파일 준비**: 예시 파일 다운로드하여 데이터 형식과 컬럼명 확인 후 자신의 농가 데이터를 동일한 형식으로 준비
3. **데이터 업로드**: 누적일사량, 외기기온 데이터가 필수로 포함된 CSV/Excel 파일 업로드 (Excel 파일의 경우 시트 선택)
4. **분석 실행**: '구간 매칭 및 비교 분석' 버튼을 클릭하여 최적 환경 조건과 비교 분석

**🎯 결과 해석**
- **노란색 행**: 최고 수확량 달성 조건
- **🟢 적정**: 현상 유지 | **🟡 미달**: 증가 필요 | **🔴 초과**: 감소 필요
- **게이지 차트**: 녹색 영역이 적정 범위
""")