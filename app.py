"""
=================================================================
전사 출근 예정 및 현황 관리 프로그램 (Streamlit 버전)
보물섬수산 HR 대시보드
=================================================================
"""

import streamlit as st
import pandas as pd
import numpy as np
from datetime import datetime
from pathlib import Path
import io
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side

st.set_page_config(
    page_title="출근현황 관리 대시보드",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded"
)

DEPARTMENTS = [
    "회계부", "회주방", "식당 홀", "식당 주방",
    "배송팀", "물류팀", "해썹가공공장", "마트 입점팀"
]
STATUS_OPTIONS = ["출근완료", "지각", "결근", "휴무"]
DATA_DIR = Path("attendance_data")
DATA_FILE = DATA_DIR / "attendance.csv"

def create_initial_data():
    np.random.seed(42)
    department_sizes = {
        "회계부": 25, "회주방": 40, "식당 홀": 45, "식당 주방": 50,
        "배송팀": 55, "물류팀": 60, "해썹가공공장": 35, "마트 입점팀": 30
    }
    data = []
    today = datetime.now().date()
    for dept, count in department_sizes.items():
        for i in range(1, count + 1):
            scheduled_hour = np.random.randint(8, 10)
            scheduled_minute = np.random.choice([0, 15, 30, 45])
            scheduled_time = f"{scheduled_hour:02d}:{scheduled_minute:02d}"
            if np.random.rand() > 0.3:
                actual_hour = np.random.randint(8, 11)
                actual_minute = np.random.choice([0, 15, 30, 45])
                actual_time = f"{actual_hour:02d}:{actual_minute:02d}"
            else:
                actual_time = ""
            if actual_time:
                status = "지각" if int(actual_time.split(":")[0]) >= 10 else "출근완료"
            else:
                r = np.random.rand()
                status = "결근" if r > 0.9 else ("휴무" if r > 0.7 else "")
            data.append({
                "날짜": today, "부서": dept,
                "이름": f"{dept}_{i:03d}",
                "출근예정시간": scheduled_time,
                "실제출근시간": actual_time,
                "근무상태": status, "비고": ""
            })
    return pd.DataFrame(data)
def load_data():
    if 'attendance_df' not in st.session_state:
        DATA_DIR.mkdir(exist_ok=True)
        if DATA_FILE.exists():
            df = pd.read_csv(DATA_FILE)
            df['날짜'] = pd.to_datetime(df['날짜']).dt.date
        else:
            df = create_initial_data()
            df.to_csv(DATA_FILE, index=False, encoding='utf-8-sig')
        st.session_state['attendance_df'] = df
    return st.session_state['attendance_df'].copy()

def save_data(df):
    DATA_DIR.mkdir(exist_ok=True)
    df.to_csv(DATA_FILE, index=False, encoding='utf-8-sig')
    st.session_state['attendance_df'] = df

def get_today_data(df):
    today = datetime.now().date()
    return df[df['날짜'] == today].copy()

def calculate_statistics(df):
    today_df = get_today_data(df)
    total = len(today_df)
    completed = len(today_df[today_df['근무상태'] == '출근완료'])
    late = len(today_df[today_df['근무상태'] == '지각'])
    absent = len(today_df[today_df['근무상태'] == '결근'])
    vacation = len(today_df[today_df['근무상태'] == '휴무'])
    working = total - vacation
    rate = round((completed / working) * 100, 1) if working > 0 else 0
    return {
        'total': total, 'completed': completed, 'late': late,
        'absent': absent, 'vacation': vacation,
        'working': working, 'rate': rate
    }

def export_to_excel(df):
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df.to_excel(writer, sheet_name='출근현황', index=False)
        ws = writer.sheets['출근현황']
        header_fill = PatternFill(start_color="4472C4", end_color="4472C4", fill_type="solid")
        header_font = Font(bold=True, color="FFFFFF", size=11)
        thin = Border(
            left=Side(style='thin'), right=Side(style='thin'),
            top=Side(style='thin'), bottom=Side(style='thin')
        )
        for cell in ws[1]:
            cell.fill = header_fill
            cell.font = header_font
            cell.alignment = Alignment(horizontal="center", vertical="center")
        for row in ws.iter_rows(min_row=2, max_row=len(df)+1):
            for cell in row:
                cell.border = thin
                cell.alignment = Alignment(horizontal="center", vertical="center")
        for col in ws.columns:
            max_len = max((len(str(c.value)) for c in col if c.value), default=0)
            ws.column_dimensions[col[0].column_letter].width = min(max_len + 2, 20)
    output.seek(0)
    return output.getvalue()
def department_manager_mode():
    st.subheader("📋 부서 관리자 모드 - 출근 현황 입력")
    today = datetime.now().strftime("%Y년 %m월 %d일 (%A)")
    st.info(f"📅 **오늘: {today}**")

    selected_dept = st.selectbox("🏢 부서를 선택하세요:", DEPARTMENTS, key="dept_select")

    df = load_data()
    today_df = get_today_data(df)
    dept_data = today_df[today_df['부서'] == selected_dept].copy()

    st.write(f"**{selected_dept}** - 총 {len(dept_data)}명")
    st.markdown("---")

    c1, c2, c3, c4 = st.columns([2, 2, 2, 2])
    with c1: st.markdown("**이름**")
    with c2: st.markdown("**출근예정**")
    with c3: st.markdown("**실제출근**")
    with c4: st.markdown("**근무상태**")
    st.divider()

    new_values = {}
    for _, row in dept_data.iterrows():
        c1, c2, c3, c4 = st.columns([2, 2, 2, 2])
        with c1: st.write(f"👤 {row['이름']}")
        with c2: st.write(row['출근예정시간'])
        with c3:
            actual = st.text_input(
                "실제출근시간",
                value=str(row['실제출근시간']) if row['실제출근시간'] else "",
                placeholder="HH:MM",
                key=f"t_{row['이름']}",
                label_visibility="collapsed"
            )
        with c4:
            idx = STATUS_OPTIONS.index(row['근무상태']) if row['근무상태'] in STATUS_OPTIONS else 0
            status = st.selectbox(
                "근무상태", STATUS_OPTIONS, index=idx,
                key=f"s_{row['이름']}",
                label_visibility="collapsed"
            )
        new_values[row['이름']] = {'actual': actual, 'status': status}

    st.markdown("---")
    _, c2, _ = st.columns([1, 1, 1])
    with c2:
        if st.button("💾 저장", use_container_width=True, type="primary"):
            df = load_data()
            today = datetime.now().date()
            for name, vals in new_values.items():
                mask = (df['날짜'] == today) & (df['이름'] == name)
                df.loc[mask, '실제출근시간'] = vals['actual']
                df.loc[mask, '근무상태'] = vals['status']
            save_data(df)
            st.success("✅ 데이터가 저장되었습니다!")
def admin_dashboard_mode():
    st.subheader("📊 총괄 관리자 모드 - 전사 출근 현황 대시보드")
    today = datetime.now().strftime("%Y년 %m월 %d일 (%A)")
    st.info(f"📅 **기준일: {today}**")

    df = load_data()
    stats = calculate_statistics(df)
    today_df = get_today_data(df)

    st.markdown("### 📈 오늘의 주요 지표")
    c1, c2, c3, c4 = st.columns(4)
    c1.metric("📊 출근율", f"{stats['rate']}%", f"({stats['completed']}/{stats['working']}명)", delta_color="off")
    c2.metric("✅ 출근완료", stats['completed'], "명", delta_color="off")
    c3.metric("⏰ 지각", stats['late'], "명", delta_color="inverse")
    c4.metric("❌ 결근", stats['absent'], "명", delta_color="inverse")

    st.divider()
    st.markdown("### 🏢 부서별 상세 현황")
    summary = []
    for dept in DEPARTMENTS:
        d = today_df[today_df['부서'] == dept]
        if len(d) == 0: continue
        comp = len(d[d['근무상태'] == '출근완료'])
        late = len(d[d['근무상태'] == '지각'])
        absent = len(d[d['근무상태'] == '결근'])
        vac = len(d[d['근무상태'] == '휴무'])
        working = len(d) - vac
        rate = round((comp / working) * 100, 1) if working > 0 else 0
        summary.append({
            '부서': dept, '총인원': len(d), '출근완료': comp,
            '지각': late, '결근': absent, '휴무': vac, '출근율': f"{rate}%"
        })
    st.dataframe(pd.DataFrame(summary), use_container_width=True, hide_index=True)

    st.divider()
    st.markdown("### 👥 전체 직원 상세 리스트")
    c1, c2 = st.columns(2)
    with c1:
        filter_dept = st.multiselect("부서 필터링:", DEPARTMENTS, key="filter_dept")
    with c2:
        filter_status = st.multiselect("근무상태 필터링:", STATUS_OPTIONS, key="filter_status")

    filtered = today_df.copy()
    if filter_dept: filtered = filtered[filtered['부서'].isin(filter_dept)]
    if filter_status: filtered = filtered[filtered['근무상태'].isin(filter_status)]
    filtered = filtered.sort_values(['부서', '이름']).reset_index(drop=True)
    st.dataframe(
        filtered[['날짜','부서','이름','출근예정시간','실제출근시간','근무상태','비고']],
        use_container_width=True, hide_index=True, height=400
    )

    st.divider()
    st.markdown("### 📥 데이터 다운로드")
    c1, c2 = st.columns(2)
    with c1:
        st.download_button(
            "📊 현황 엑셀로 다운로드",
            export_to_excel(filtered),
            f"출근현황_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True
        )
    with c2:
        if st.button("🔄 데이터 새로고침", use_container_width=True):
            st.rerun()


def main():
    load_data()
    st.title("🏢 보물섬수산 출근현황 관리 대시보드")
    st.sidebar.markdown("---")
    st.sidebar.title("🎯 메뉴")
    mode = st.sidebar.radio(
        "작업 모드를 선택하세요:",
        ["📋 부서 관리자 모드", "📊 총괄 관리자 모드"]
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

    if "부서 관리자" in mode:
        department_manager_mode()
    else:
        admin_dashboard_mode()


if __name__ == "__main__":
    main()