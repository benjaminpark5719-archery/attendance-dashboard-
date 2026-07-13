"""
=================================================================
전사 출근 예정 및 현황 관리 프로그램 (Streamlit 버전)
보물섬수산 HR 대시보드
=================================================================
기능: 부서별 출근 현황 입력 + 전사 통계 대시보드
저장: CSV 파일 (로컬 저장, 영속성 보장)
배포: Streamlit Cloud (무료 호스팅)
=================================================================
"""

import streamlit as st
import pandas as pd
import numpy as np
from datetime import datetime, timedelta
import os
from pathlib import Path
import io
from openpyxl import Workbook
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
from openpyxl.utils.dataframe import dataframe_to_rows

# ========================================
# 1️⃣ 페이지 설정
# ========================================
st.set_page_config(
    page_title="출근현황 관리 대시보드",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded"
)

# 커스텀 CSS (모바일 반응형)
st.markdown("""
    <style>
    [data-testid="stMetricValue"] {
        font-size: 2.5rem;
    }
    .stTabs [data-baseweb="tab-list"] {
        gap: 2rem;
    }
    @media (max-width: 640px) {
        .stTabs [data-baseweb="tab-list"] {
            gap: 0.5rem;
        }
    }
    </style>
    """, unsafe_allow_html=True)

# ========================================
# 2️⃣ 데이터 저장소 설정
# ========================================

# 데이터 파일 경로
DATA_DIR = Path("attendance_data")
DATA_DIR.mkdir(exist_ok=True)
DATA_FILE = DATA_DIR / "attendance.csv"

# 부서 리스트
DEPARTMENTS = [
    "회계부", "회주방", "식당 홀", "식당 주방",
    "배송팀", "물류팀", "해썹가공공장", "마트 입점팀"
]

# 근무 상태 옵션
STATUS_OPTIONS = ["출근완료", "지각", "결근", "휴무"]

# ========================================
# 3️⃣ 데이터 로드/초기화 함수
# ========================================

@st.cache_resource
def initialize_session_state():
    """Streamlit 세션 상태 초기화"""
    if 'attendance_df' not in st.session_state:
        st.session_state.attendance_df = load_or_create_data()

def load_or_create_data():
    """
    CSV 파일에서 데이터 로드
    없으면 초기 데이터 생성
    """
    if DATA_FILE.exists():
        df = pd.read_csv(DATA_FILE)
        # 날짜 컬럼을 datetime으로 변환
        df['날짜'] = pd.to_datetime(df['날짜']).dt.date
        return df
    else:
        # 초기 데이터 생성 (약 300명, 8개 부서)
        return create_initial_data()

def create_initial_data():
    """
    초기 샘플 데이터 생성
    부서별 직원 수를 나누어 배치
    """
    np.random.seed(42)
    
    # 부서별 직원 수
    department_sizes = {
        "회계부": 25,
        "회주방": 40,
        "식당 홀": 45,
        "식당 주방": 50,
        "배송팀": 55,
        "물류팀": 60,
        "해썹가공공장": 35,
        "마트 입점팀": 30
    }
    
    data = []
    today = datetime.now().date()
    
    for dept, count in department_sizes.items():
        for i in range(1, count + 1):
            # 직원 이름 생성
            employee_name = f"{dept}_{i:03d}"
            
            # 출근 예정 시간 (08:00 ~ 10:00 사이)
            scheduled_hour = np.random.randint(8, 10)
            scheduled_minute = np.random.choice([0, 15, 30, 45])
            scheduled_time = f"{scheduled_hour:02d}:{scheduled_minute:02d}"
            
            # 실제 출근 시간 (일부는 아직 미입력)
            if np.random.rand() > 0.3:  # 70% 확률로 입력됨
                actual_hour = np.random.randint(8, 11)
                actual_minute = np.random.choice([0, 15, 30, 45])
                actual_time = f"{actual_hour:02d}:{actual_minute:02d}"
            else:
                actual_time = ""
            
            # 근무 상태 결정 (실제 시간 입력 여부에 따라)
            if actual_time:
                if int(actual_time.split(":")[0]) > 9 or \
                   (int(actual_time.split(":")[0]) == 9 and int(actual_time.split(":")[1]) > 0):
                    status = "지각"
                else:
                    status = "출근완료"
            else:
                if np.random.rand() > 0.9:
                    status = "결근"
                elif np.random.rand() > 0.7:
                    status = "휴무"
                else:
                    status = ""
            
            data.append({
                "날짜": today,
                "부서": dept,
                "이름": employee_name,
                "출근예정시간": scheduled_time,
                "실제출근시간": actual_time,
                "근무상태": status,
                "비고": ""
            })
    
    return pd.DataFrame(data)

def save_data(df):
    """데이터를 CSV 파일에 저장"""
    df.to_csv(DATA_FILE, index=False, encoding='utf-8-sig')
    st.session_state.attendance_df = df

def load_data():
    """세션 상태에서 현재 데이터 로드"""
    return st.session_state.attendance_df.copy()

# ========================================
# 4️⃣ 유틸리티 함수
# ========================================

def get_today_data(df):
    """오늘 날짜의 데이터만 필터링"""
    today = datetime.now().date()
    return df[df['날짜'] == today].copy()

def get_department_employees(df, department):
    """특정 부서의 직원 리스트"""
    return df[df['부서'] == department].copy()

def calculate_statistics(df):
    """통계 계산"""
    today_df = get_today_data(df)
    
    total_employees = len(today_df)
    completed = len(today_df[today_df['근무상태'] == '출근완료'])
    late = len(today_df[today_df['근무상태'] == '지각'])
    absent = len(today_df[today_df['근무상태'] == '결근'])
    vacation = len(today_df[today_df['근무상태'] == '휴무'])
    
    # 출근 완료율 (휴무 제외)
    working_employees = total_employees - vacation
    if working_employees > 0:
        attendance_rate = round((completed / working_employees) * 100, 1)
    else:
        attendance_rate = 0
    
    return {
        'total': total_employees,
        'completed': completed,
        'late': late,
        'absent': absent,
        'vacation': vacation,
        'working': working_employees,
        'attendance_rate': attendance_rate
    }

def export_to_excel(df):
    """DataFrame을 Excel 파일로 변환 (바이너리)"""
    output = io.BytesIO()
    
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df.to_excel(writer, sheet_name='출근현황', index=False)
        
        # 스타일 지정
        workbook = writer.book
        worksheet = writer.sheets['출근현황']
        
        # 헤더 스타일
        header_fill = PatternFill(start_color="4472C4", end_color="4472C4", fill_type="solid")
        header_font = Font(bold=True, color="FFFFFF", size=11)
        
        for cell in worksheet[1]:
            cell.fill = header_fill
            cell.font = header_font
            cell.alignment = Alignment(horizontal="center", vertical="center")
        
        # 데이터 셀 스타일
        thin_border = Border(
            left=Side(style='thin'),
            right=Side(style='thin'),
            top=Side(style='thin'),
            bottom=Side(style='thin')
        )
        
        for row in worksheet.iter_rows(min_row=2, max_row=len(df)+1):
            for cell in row:
                cell.border = thin_border
                cell.alignment = Alignment(horizontal="center", vertical="center")
        
        # 컬럼 너비 자동 조정
        for column in worksheet.columns:
            max_length = 0
            column_letter = column[0].column_letter
            for cell in column:
                try:
                    if len(str(cell.value)) > max_length:
                        max_length = len(str(cell.value))
                except:
                    pass
            adjusted_width = min(max_length + 2, 20)
            worksheet.column_dimensions[column_letter].width = adjusted_width
    
    output.seek(0)
    return output.getvalue()

# ========================================
# 5️⃣ 부서 관리자 모드 UI
# ========================================

def department_manager_mode():
    """부서 관리자 화면"""
    st.subheader("📋 부서 관리자 모드 - 출근 현황 입력")
    
    # 오늘 날짜 표시
    today = datetime.now().strftime("%Y년 %m월 %d일 (%A)")
    st.info(f"📅 **오늘: {today}**")
    
    # 세션 상태 초기화
    if 'edit_mode' not in st.session_state:
        st.session_state.edit_mode = {}
    
    # 부서 선택
    selected_dept = st.selectbox(
        "🏢 부서를 선택하세요:",
        DEPARTMENTS,
        key="dept_select"
    )
    
    # 선택된 부서의 직원 데이터
    df = load_data()
    dept_data = get_department_employees(get_today_data(df), selected_dept)
    
    st.write(f"**{selected_dept}** - 총 {len(dept_data)}명")
    
    # 직원별 입력 폼
    st.markdown("---")
    
    # 데이터 편집용 딕셔너리 (세션 상태에 저장)
    if selected_dept not in st.session_state.edit_mode:
        st.session_state.edit_mode[selected_dept] = {}
    
    edit_data = st.session_state.edit_mode[selected_dept]
    
    # 직원 정보 표시 및 입력
    col1, col2, col3, col4 = st.columns([2, 2, 2, 2])
    
    with col1:
        st.markdown("**이름**")
    with col2:
        st.markdown("**출근예정**")
    with col3:
        st.markdown("**실제출근**")
    with col4:
        st.markdown("**근무상태**")
    
    st.divider()
    
    # 모바일 대응: 스택 레이아웃
    for idx, row in dept_data.iterrows():
        col1, col2, col3, col4 = st.columns([2, 2, 2, 2])
        
        employee_id = f"{selected_dept}_{row['이름']}"
        
        with col1:
            st.write(f"👤 {row['이름']}")
        
        with col2:
            st.write(row['출근예정시간'])
        
        with col3:
            # 실제 출근 시간 입력
            actual_time = st.text_input(
                label="실제출근시간",
                value=edit_data.get(employee_id, {}).get('actual_time', row['실제출근시간']),
                placeholder="HH:MM",
                key=f"actual_{employee_id}",
                label_visibility="collapsed"
            )
            edit_data[employee_id] = {**edit_data.get(employee_id, {}), 'actual_time': actual_time}
        
        with col4:
            # 근무 상태 선택
            status = st.selectbox(
                label="근무상태",
                options=STATUS_OPTIONS,
                index=STATUS_OPTIONS.index(row['근무상태']) if row['근무상태'] in STATUS_OPTIONS else 0,
                key=f"status_{employee_id}",
                label_visibility="collapsed"
            )
            edit_data[employee_id] = {**edit_data.get(employee_id, {}), 'status': status}
    
    st.markdown("---")
    
    # 저장 버튼
    col1, col2, col3 = st.columns([1, 1, 1])
    
    with col2:
        if st.button("💾 저장", use_container_width=True, type="primary"):
            # 입력 데이터를 메인 DataFrame에 반영
            df = load_data()
            today = datetime.now().date()
            
            for employee_id, edited_values in edit_data.items():
                mask = (df['날짜'] == today) & (df['이름'] == employee_id.split('_', 1)[1])
                if mask.any():
                    df.loc[mask, '실제출근시간'] = edited_values.get('actual_time', '')
                    df.loc[mask, '근무상태'] = edited_values.get('status', '')
            
            # 데이터 저장
            save_data(df)
            st.success("✅ 데이터가 저장되었습니다!")
            st.session_state.edit_mode[selected_dept] = {}

# ========================================
# 6️⃣ 총괄 관리자 모드 UI
# ========================================

def admin_dashboard_mode():
    """총괄 관리자 대시보드"""
    st.subheader("📊 총괄 관리자 모드 - 전사 출근 현황 대시보드")
    
    # 오늘 날짜 표시
    today = datetime.now().strftime("%Y년 %m월 %d일 (%A)")
    st.info(f"📅 **기준일: {today}**")
    
    # 데이터 로드
    df = load_data()
    today_df = get_today_data(df)
    stats = calculate_statistics(df)
    
    # ===== 상단 통계 카드 =====
    st.markdown("### 📈 오늘의 주요 지표")
    
    col1, col2, col3, col4 = st.columns(4)
    
    with col1:
        st.metric(
            label="📊 출근율",
            value=f"{stats['attendance_rate']}%",
            delta=f"({stats['completed']}/{stats['working']}명)",
            delta_color="off"
        )
    
    with col2:
        st.metric(
            label="✅ 출근완료",
            value=stats['completed'],
            delta="명",
            delta_color="off"
        )
    
    with col3:
        st.metric(
            label="⏰ 지각",
            value=stats['late'],
            delta="명",
            delta_color="inverse"
        )
    
    with col4:
        st.metric(
            label="❌ 결근",
            value=stats['absent'],
            delta="명",
            delta_color="inverse"
        )
    
    st.divider()
    
    # ===== 부서별 현황표 =====
    st.markdown("### 🏢 부서별 상세 현황")
    
    department_summary = []
    
    for dept in DEPARTMENTS:
        dept_data = get_department_employees(today_df, dept)
        if len(dept_data) > 0:
            dept_completed = len(dept_data[dept_data['근무상태'] == '출근완료'])
            dept_late = len(dept_data[dept_data['근무상태'] == '지각'])
            dept_absent = len(dept_data[dept_data['근무상태'] == '결근'])
            dept_vacation = len(dept_data[dept_data['근무상태'] == '휴무'])
            dept_working = len(dept_data) - dept_vacation
            
            if dept_working > 0:
                dept_rate = round((dept_completed / dept_working) * 100, 1)
            else:
                dept_rate = 0
            
            department_summary.append({
                '부서': dept,
                '총인원': len(dept_data),
                '출근완료': dept_completed,
                '지각': dept_late,
                '결근': dept_absent,
                '휴무': dept_vacation,
                '출근율': f"{dept_rate}%"
            })
    
    summary_df = pd.DataFrame(department_summary)
    
    # 테이블 표시 (st.dataframe 스타일링)
    st.dataframe(
        summary_df,
        use_container_width=True,
        hide_index=True,
        column_config={
            '부서': st.column_config.TextColumn(width="medium"),
            '총인원': st.column_config.NumberColumn(width="small"),
            '출근완료': st.column_config.NumberColumn(width="small"),
            '지각': st.column_config.NumberColumn(width="small"),
            '결근': st.column_config.NumberColumn(width="small"),
            '휴무': st.column_config.NumberColumn(width="small"),
            '출근율': st.column_config.TextColumn(width="small")
        }
    )
    
    st.divider()
    
    # ===== 전체 직원 현황표 =====
    st.markdown("### 👥 전체 직원 상세 리스트")
    
    # 필터 옵션
    col1, col2 = st.columns(2)
    
    with col1:
        filter_dept = st.multiselect(
            "부서 필터링 (선택 안 하면 전체):",
            DEPARTMENTS,
            default=None,
            key="filter_dept"
        )
    
    with col2:
        filter_status = st.multiselect(
            "근무상태 필터링 (선택 안 하면 전체):",
            STATUS_OPTIONS,
            default=None,
            key="filter_status"
        )
    
    # 필터 적용
    filtered_df = today_df.copy()
    
    if filter_dept:
        filtered_df = filtered_df[filtered_df['부서'].isin(filter_dept)]
    
    if filter_status:
        filtered_df = filtered_df[filtered_df['근무상태'].isin(filter_status)]
    
    # 정렬
    filtered_df = filtered_df.sort_values(['부서', '이름']).reset_index(drop=True)
    
    # 테이블 표시
    st.dataframe(
        filtered_df[['날짜', '부서', '이름', '출근예정시간', '실제출근시간', '근무상태', '비고']],
        use_container_width=True,
        hide_index=True,
        height=400
    )
    
    st.divider()
    
    # ===== 엑셀 다운로드 =====
    st.markdown("### 📥 데이터 다운로드")
    
    excel_data = export_to_excel(filtered_df)
    
    col1, col2 = st.columns([1, 1])
    
    with col1:
        st.download_button(
            label="📊 현황 엑셀로 다운로드",
            data=excel_data,
            file_name=f"출근현황_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True
        )
    
    with col2:
        if st.button("🔄 데이터 새로고침", use_container_width=True):
            st.rerun()

# ========================================
# 7️⃣ 메인 앱 로직
# ========================================

def main():
    # 세션 상태 초기화
    initialize_session_state()
    
    # 헤더
    st.title("🏢 보물섬수산 출근현황 관리 대시보드")
    
    # 사이드바 모드 선택
    st.sidebar.markdown("---")
    st.sidebar.title("🎯 메뉴")
    
    mode = st.sidebar.radio(
        "작업 모드를 선택하세요:",
        options=["📋 부서 관리자 모드", "📊 총괄 관리자 모드"],
        index=0
    )
    
    st.sidebar.markdown("---")
    st.sidebar.markdown("""
    ### 💡 사용 가이드
    
    **부서 관리자:**
    - 자신의 부서를 선택
    - 직원별 실제 출근 시간 입력
    - 근무 상태 선택
    - 저장 버튼 클릭
    
    **총괄 관리자:**
    - 전사 출근율 한눈에 파악
    - 부서별 상세 현황 확인
    - 필터링해서 리스트 확인
    - 엑셀로 다운로드
    """)
    
    st.sidebar.markdown("---")
    st.sidebar.caption("v1.0 | 2025년 보물섬수산 HR시스템")
    
    # 모드별 화면 표시
    if "부서 관리자" in mode:
        department_manager_mode()
    else:
        admin_dashboard_mode()

if __name__ == "__main__":
    main()
