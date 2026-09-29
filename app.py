# -*- coding: utf-8 -*-
"""
================================================================================
보물섬수산 HR 시스템 — 통합 app.py
================================================================================
"""
import io
import base64
from datetime import date, datetime, time, timedelta
import pandas as pd
import streamlit as st
from PIL import Image

try:
    from supabase import create_client
    _HAS_SUPABASE = True
except Exception:
    _HAS_SUPABASE = False

try:
    from streamlit_option_menu import option_menu
    _HAS_OPTION_MENU = True
except Exception:
    _HAS_OPTION_MENU = False

try:
    from openpyxl import Workbook
    from openpyxl.worksheet.datavalidation import DataValidation
    from openpyxl.styles import Font, PatternFill, Alignment
    _HAS_OPENPYXL = True
except Exception:
    _HAS_OPENPYXL = False

try:
    from fpdf import FPDF
    _HAS_FPDF = True
except Exception:
    _HAS_FPDF = False

try:
    import holidays as _kr_holidays
    _HAS_HOLIDAYS = True
except Exception:
    _HAS_HOLIDAYS = False

st.set_page_config(page_title="보물섬수산 HR", page_icon="🐟", layout="wide")
APP_TITLE = "보물섬수산"
APP_SUBTITLE = "HR 통합 관리 시스템"
LOGO_EMOJI = "🐟"

LOGIN_ID = "admin"
LOGIN_PW = "admin1234!"

DEFAULT_MENU = [
    ("전사현황", "📊"),
    ("출근입력", "📋"),
    ("근무계획", "🗓️"),
    ("이력조회", "🔍"),
    ("부서관리", "🏢"),
    ("직원관리", "👤"),
    ("TBM안전관리", "🦺"),
    ("급여관리", "💰"),
    ("시스템관리", "⚙️"),
]

TB = {
    "departments": "departments",
    "employees": "employees",
    "attendance": "attendance",
    "shift_plans": "shift_plans",
    "tbm_logs": "tbm_logs",
}

SHIFT_SCHEDULE = {
    "오전": {"start": time(6, 0), "end": time(15, 0), "color": "#FFF3CD"},
    "오후": {"start": time(13, 0), "end": time(22, 0), "color": "#D1E7DD"},
    "야간": {"start": time(22, 0), "end": time(6, 0), "color": "#CFE2FF"},
    "휴무": {"start": None, "end": None, "color": "#E2E3E5"},
}
SHIFT_OPTIONS = list(SHIFT_SCHEDULE.keys())
GRACE_MINUTES = 0
OVERTIME_RATE = 1.5
HOLIDAY_RATE = 1.5
STATUS_COLOR = {"정상": "#198754", "지각": "#dc3545", "결근": "#6c757d"}
FONT_PATH = "fonts/NanumGothic.ttf"

CUSTOM_CSS = """
<style>
:root { --brand:#0d6efd; --brand-dark:#0a58ca; }
section[data-testid="stSidebar"] { background:#0b1f33; }
section[data-testid="stSidebar"] * { color:#e9eef5; }
.brand-head { display:flex; align-items:center; gap:.55rem; padding:.4rem .2rem 1rem .2rem; border-bottom:1px solid rgba(255,255,255,.12); margin-bottom:.8rem; }
.brand-emoji { font-size:1.9rem; }
.brand-name  { font-size:1.15rem; font-weight:800; line-height:1.1; }
.brand-sub   { font-size:.72rem; opacity:.75; }
.metric-card { background:#fff; border:1px solid #eef0f3; border-radius:16px; padding:.9rem 1rem; box-shadow:0 2px 10px rgba(16,24,40,.05); }
.metric-card .label { font-size:.8rem; color:#667085; margin-bottom:.15rem; }
.metric-card .value { font-size:1.7rem; font-weight:800; line-height:1; }
.metric-card .delta { font-size:.72rem; color:#98a2b3; }
.badge { display:inline-block; padding:.18rem .55rem; border-radius:999px; font-size:.72rem; font-weight:700; color:#fff; }
.block-container { padding-top:3.5rem !important; }
.main h1, .main h2, .main h3, div[data-testid="stHeading"] { line-height:1.4; margin-top:.2rem; }
div[data-testid="stDataFrame"] { border-radius:12px; overflow:hidden; }
@media (max-width:640px){.metric-card .value { font-size:1.35rem; } .block-container { padding-left:.6rem; padding-right:.6rem; } }
</style>
"""

def badge(text, color):
    return f'<span class="badge" style="background:{color}">{text}</span>'

def metric_card(label, value, delta=""):
    st.markdown(
        f'<div class="metric-card"><div class="label">{label}</div>'
        f'<div class="value">{value}</div><div class="delta">{delta}</div></div>',
        unsafe_allow_html=True,
    )

@st.cache_resource(show_spinner=False)
def get_supabase():
    if not _HAS_SUPABASE:
        return None
    try:
        url = st.secrets["SUPABASE_URL"]
        key = st.secrets["SUPABASE_KEY"]
    except Exception:
        return None
    if not url or not key:
        return None
    try:
        return create_client(url, key)
    except Exception:
        return None

SB = get_supabase()
USE_SUPABASE = SB is not None

def _seed_demo():
    if st.session_state.get("_demo_ready"):
        return
    today = date.today()
    st.session_state.demo_departments = [
        {"id": 1, "name": "영업부", "sort_order": 1},
        {"id": 2, "name": "가공1팀", "sort_order": 2},
        {"id": 3, "name": "물류센터", "sort_order": 3},
        {"id": 4, "name": "매장운영", "sort_order": 4},
    ]
    st.session_state.demo_employees = [
        {"id": 1, "emp_no": "B001", "name": "김민수", "department": "영업부", "position": "과장", "hire_date": "2021-03-02", "phone": "010-1111-2222", "health_cert_expiry": str(today + timedelta(days=12)), "hourly_wage": 12000, "status": "재직"},
        {"id": 2, "emp_no": "B002", "name": "이서연", "department": "가공1팀", "position": "사원", "hire_date": "2023-06-01", "phone": "010-3333-4444", "health_cert_expiry": str(today + timedelta(days=45)), "hourly_wage": 10500, "status": "재직"},
        {"id": 3, "emp_no": "B003", "name": "박준호", "department": "물류센터", "position": "대리", "hire_date": "2022-01-10", "phone": "010-5555-6666", "health_cert_expiry": str(today - timedelta(days=3)), "hourly_wage": 11000, "status": "재직"},
        {"id": 4, "emp_no": "B004", "name": "최지우", "department": "매장운영", "position": "사원", "hire_date": "2024-02-19", "phone": "010-7777-8888", "health_cert_expiry": str(today + timedelta(days=120)), "hourly_wage": 10200, "status": "재직"},
    ]
    st.session_state.demo_attendance = []
    st.session_state.demo_shift_plans = []
    st.session_state.demo_tbm_logs = []
    st.session_state._demo_ready = True

def _demo_table(name):
    _seed_demo()
    return st.session_state[f"demo_{name}"]

def db_select(name):
    if USE_SUPABASE:
        try:
            res = SB.table(TB[name]).select("*").execute()
            return pd.DataFrame(res.data or [])
        except Exception as e:
            st.session_state.setdefault("_db_errors", []).append(f"{name} select: {e}")
            return pd.DataFrame(_demo_table(name))
    return pd.DataFrame(_demo_table(name))

def db_insert(name, row: dict):
    if USE_SUPABASE:
        try:
            SB.table(TB[name]).insert(row).execute()
            return True
        except Exception as e:
            st.warning(f"저장 실패({name}): {e}")
            return False
    tbl = _demo_table(name)
    row = dict(row)
    row.setdefault("id", (max([r.get("id", 0) for r in tbl], default=0) + 1))
    tbl.append(row)
    return True

def db_delete(name, id_val):
    if USE_SUPABASE:
        try:
            SB.table(TB[name]).delete().eq("id", id_val).execute()
            return True
        except Exception as e:
            st.warning(f"삭제 실패({name}): {e}")
            return False
    tbl = _demo_table(name)
    st.session_state[f"demo_{name}"] = [r for r in tbl if r.get("id") != id_val]
    return True

def db_update(name, id_val, patch: dict):
    if USE_SUPABASE:
        try:
            SB.table(TB[name]).update(patch).eq("id", id_val).execute()
            return True
        except Exception as e:
            st.warning(f"수정 실패({name}): {e}")
            return False
    for r in _demo_table(name):
        if r.get("id") == id_val:
            r.update(patch)
    return True

def db_upsert_attendance(row: dict):
    key_emp, key_date = row.get("employee_id"), row.get("work_date")
    if USE_SUPABASE:
        try:
            SB.table(TB["attendance"]).delete().eq("employee_id", key_emp).eq("work_date", key_date).execute()
            SB.table(TB["attendance"]).insert(row).execute()
            return True
        except Exception as e:
            st.warning(f"출근 저장 실패: {e}")
            return False
    tbl = _demo_table("attendance")
    st.session_state["demo_attendance"] = [r for r in tbl if not (r.get("employee_id") == key_emp and r.get("work_date") == key_date)]
    return db_insert("attendance", row)

def compress_image(uploaded_file, max_side=1280, quality=70):
    img = Image.open(uploaded_file)
    if img.mode in ("RGBA", "P", "LA"):
        img = img.convert("RGB")
    img.thumbnail((max_side, max_side))
    buf = io.BytesIO()
    img.save(buf, format="JPEG", quality=quality, optimize=True)
    return buf.getvalue()

def _parse_time(v):
    if v in (None, "", "None"):
        return None
    if isinstance(v, time):
        return v
    for fmt in ("%H:%M:%S", "%H:%M"):
        try:
            return datetime.strptime(str(v), fmt).time()
        except ValueError:
            continue
    return None

def judge_attendance(planned_start, actual_start, grace=GRACE_MINUTES):
    planned_start = _parse_time(planned_start)
    actual_start = _parse_time(actual_start)
    if planned_start is None:
        return "정상" if actual_start else "결근"
    if actual_start is None:
        return "결근"
    limit = (datetime.combine(date.today(), planned_start) + timedelta(minutes=grace)).time()
    return "지각" if actual_start > limit else "정상"

def get_kr_holidays(year):
    if _HAS_HOLIDAYS:
        try:
            return set(_kr_holidays.KR(years=year).keys())
        except Exception:
            pass
    return set()

def build_shift_template_xlsx(employees_df):
    if not _HAS_OPENPYXL:
        return None
    wb = Workbook()
    ws = wb.active
    ws.title = "근무계획"
    headers = ["부서", "이름", "월", "화", "수", "목", "금", "토", "일"]
    ws.append(headers)
    head_fill = PatternFill("solid", fgColor="0D6EFD")
    for c in range(1, len(headers) + 1):
        cell = ws.cell(row=1, column=c)
        cell.font = Font(bold=True, color="FFFFFF")
        cell.fill = head_fill
        cell.alignment = Alignment(horizontal="center")
    for _, e in employees_df.iterrows():
        ws.append([e.get("department", ""), e.get("name", "")] + ["휴무"] * 7)
    dv = DataValidation(type="list", formula1='"오전,오후,야간,휴무"', allow_blank=True)
    ws.add_data_validation(dv)
    last = ws.max_row if ws.max_row > 1 else 200
    dv.add(f"C2:I{max(last, 200)}")
    for col, w in zip("ABCDEFGHI", [12, 12, 8, 8, 8, 8, 8, 8, 8]):
        ws.column_dimensions[col].width = w
    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()

def page_dashboard():
    st.title("📊 전사 출근 현황 (실시간)")
    st.write("오늘 부서별, 직원별 출근 현황 및 요약 수치를 확인합니다.")
    emp_df = db_select("employees")
    att_df = db_select("attendance")
    today_str = str(date.today())
    today_att = att_df[att_df["work_date"] == today_str] if not att_df.empty else pd.DataFrame()
    
    total_emp = len(emp_df[emp_df["status"] == "재직"]) if not emp_df.empty else 0
    present_cnt = len(today_att[today_att["status"] == "정상"]) if not today_att.empty else 0
    late_cnt = len(today_att[today_att["status"] == "지각"]) if not today_att.empty else 0
    absent_cnt = max(0, total_emp - (present_cnt + late_cnt))

    c1, c2, c3, c4 = st.columns(4)
    with c1: metric_card("전체 재직원", f"{total_emp} 명")
    with c2: metric_card("정상 출근", f"{present_cnt} 명")
    with c3: metric_card("지각", f"{late_cnt} 명")
    with c4: metric_card("미출근/결근", f"{absent_cnt} 명")

def page_clockin():
    st.subheader("📋 출근 입력 (현장용)")
    emp_df = db_select("employees")
    if emp_df.empty:
        st.info("등록된 직원이 없습니다.")
        return
    active_emp = emp_df[emp_df["status"] == "재직"]
    dept_list = sorted(active_emp["department"].dropna().unique().tolist())
    sel_dept = st.selectbox("부서 선택", ["전체"] + dept_list)
    filtered = active_emp if sel_dept == "전체" else active_emp[active_emp["department"] == sel_dept]
    sel_emp_name = st.selectbox("직원 선택", filtered["name"].tolist())
    target = filtered[filtered["name"] == sel_emp_name].iloc[0]

    with st.form("clockin_form"):
        st.write(f"👤 **{target['name']}** ({target['department']}) 님 출근 입력")
        now_time = datetime.now().time()
        in_time = st.time_input("실제 출근 시간", value=now_time)
        if st.form_submit_button("🚀 출근 완료 등록", type="primary"):
            row = {
                "employee_id": target["id"],
                "emp_no": target["emp_no"],
                "name": target["name"],
                "department": target["department"],
                "work_date": str(date.today()),
                "actual_start": str(in_time),
                "status": "정상"
            }
            if db_upsert_attendance(row):
                st.balloons()
                st.success("출근 기록이 성공적으로 저장되었습니다!")

def page_schedule():
    st.subheader("📅 주간 근무계획 (순환 탄력근무제)")
    emp_df = db_select("employees")
    if not emp_df.empty:
        xlsx_data = build_shift_template_xlsx(emp_df)
        if xlsx_data:
            st.download_button(
                label="📥 주간 순환근무 표준 엑셀 양식 다운로드",
                data=xlsx_data,
                file_name="보물섬수산_주간근무계획_양식.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
            
    selected_monday = st.date_input("계획 주 (월요일 기준)", value=date.today() - timedelta(days=date.today().weekday()))
    
    tab1, tab2 = st.tabs(["📁 엑셀 일괄 업로드", "✏️ 직접 등록/수정"])
    
    with tab1:
        st.caption("작성한 근무계획 엑셀 파일(.xlsx)을 업로드하세요.")
        uploaded_file = st.file_uploader("엑셀 파일 선택", type=["xlsx", "xls"], key="sched_upload")
        if uploaded_file is not None:
            try:
                df_raw = pd.read_excel(uploaded_file)
                st.write("▼ 업로드 데이터 미리보기")
                st.dataframe(df_raw.head(5), use_container_width=True)
                if st.button("🚀 DB 일괄 저장 실행", type="primary"):
                    cols = [c for c in df_raw.columns if any(day in str(c) for day in ["월", "화", "수", "목", "금", "토", "일"])]
                    name_col = [c for c in df_raw.columns if "이름" in str(c) or "성명" in str(c)]
                    name_field = name_col[0] if name_col else df_raw.columns[1]
                    
                    success_cnt = 0
                    for idx, row in df_raw.iterrows():
                        emp_name = str(row[name_field]).strip()
                        if not emp_name or emp_name == "nan": continue
                        matched_emp = emp_df[emp_df["name"] == emp_name]
                        emp_no = matched_emp.iloc[0]["emp_no"] if not matched_emp.empty else ""
                        dept = matched_emp.iloc[0]["department"] if not matched_emp.empty else row.get("부서", "")
                        plan_row = {"week_start": str(selected_monday), "emp_no": emp_no, "name": emp_name, "department": dept}
                        days_keys = ["mon", "tue", "wed", "thu", "fri", "sat", "sun"]
                        for d_idx, day_col in enumerate(cols[:7]):
                            val = str(row[day_col]).strip() if pd.notna(row[day_col]) else "휴무"
                            plan_row[days_keys[d_idx]] = val
                        if db_insert("shift_plans", plan_row): success_cnt += 1
                    st.balloons()
                    st.success(f"🎉 성공! 총 {success_cnt}명의 주간 스케줄이 DB에 등록되었습니다.")
                    st.rerun()
            except Exception as e:
                st.error(f"엑셀 처리 중 오류가 발생했습니다: {e}")

    with tab2:
        st.caption("직원을 선택하여 이번 주 요일별 근무조(오전/오후/야간/휴무)를 직접 입력합니다.")
        if emp_df.empty:
            st.warning("등록된 직원이 없습니다.")
        else:
            dept_list = sorted(emp_df["department"].dropna().unique().tolist())
            sel_dept = st.selectbox("부서 선택", ["전체"] + dept_list, key="sched_tab2_dept")
            filtered_emp = emp_df if sel_dept == "전체" else emp_df[emp_df["department"] == sel_dept]
            sel_emp_name = st.selectbox("직원 선택", filtered_emp["name"].tolist(), key="sched_tab2_emp")
            target_emp = filtered_emp[filtered_emp["name"] == sel_emp_name].iloc[0]
            
            existing_plans = db_select("shift_plans")
            curr_plan = {}
            if not existing_plans.empty:
                match = existing_plans[(existing_plans["emp_no"] == target_emp["emp_no"]) & (existing_plans["week_start"] == str(selected_monday))]
                if not match.empty: curr_plan = match.iloc[0].to_dict()

            with st.form("direct_schedule_form"):
                st.markdown(f"**👤 {target_emp['name']} ({target_emp['department']})** 님 주간 근무 설정")
                c_mon, c_tue, c_wed, c_thu, c_fri, c_sat, c_sun = st.columns(7)
                opts = ["휴무", "오전", "오후", "야간"]
                d_mon = c_mon.selectbox("월", opts, index=opts.index(curr_plan.get("mon", "오전")) if curr_plan.get("mon") in opts else 1)
                d_tue = c_tue.selectbox("화", opts, index=opts.index(curr_plan.get("tue", "오전")) if curr_plan.get("tue") in opts else 1)
                d_wed = c_wed.selectbox("수", opts, index=opts.index(curr_plan.get("wed", "오전")) if curr_plan.get("wed") in opts else 1)
                d_thu = c_thu.selectbox("목", opts, index=opts.index(curr_plan.get("thu", "오전")) if curr_plan.get("thu") in opts else 1)
                d_fri = c_fri.selectbox("금", opts, index=opts.index(curr_plan.get("fri", "오전")) if curr_plan.get("fri") in opts else 1)
                d_sat = c_sat.selectbox("토", opts, index=opts.index(curr_plan.get("sat", "휴무")) if curr_plan.get("sat") in opts else 0)
                d_sun = c_sun.selectbox("일", opts, index=opts.index(curr_plan.get("sun", "휴무")) if curr_plan.get("sun") in opts else 0)
                
                if st.form_submit_button("💾 근무계획 저장", type="primary", use_container_width=True):
                    save_data = {
                        "week_start": str(selected_monday),
                        "emp_no": target_emp["emp_no"],
                        "name": target_emp["name"],
                        "department": target_emp["department"],
                        "mon": d_mon, "tue": d_tue, "wed": d_wed, "thu": d_thu,
                        "fri": d_fri, "sat": d_sat, "sun": d_sun
                    }
                    if db_insert("shift_plans", save_data):
                        st.success(f"✅ {target_emp['name']} 님의 주간 근무계획이 저장되었습니다!")
                        st.rerun()

    st.markdown("### 주간 순환 근무 달력")
    plans_df = db_select("shift_plans")
    if plans_df.empty:
        st.info("해당 주에 등록된 근무계획이 없습니다.")
    else:
        week_plans = plans_df[plans_df["week_start"] == str(selected_monday)]
        if week_plans.empty:
            st.info("해당 주에 등록된 근무계획이 없습니다.")
        else:
            view_df = week_plans[["department", "name", "mon", "tue", "wed", "thu", "fri", "sat", "sun"]].copy()
            view_df.columns = ["부서", "이름", "월", "화", "수", "목", "금", "토", "일"]
            st.dataframe(view_df, use_container_width=True, hide_index=True)

def page_history():
    st.subheader("🔍 출퇴근 이력 조회")
    att_df = db_select("attendance")
    if att_df.empty:
        st.info("조회할 출퇴근 기록이 없습니다.")
        return
    st.dataframe(att_df, use_container_width=True)

def page_departments():
    st.subheader("🏢 부서 관리")
    dept_df = db_select("departments")
    st.dataframe(dept_df, use_container_width=True)

def page_employees():
    st.subheader("👤 직원 관리")
    emp_df = db_select("employees")
    st.dataframe(emp_df, use_container_width=True)

def page_tbm():
    st.subheader("🦺 TBM 안전보건일지")
    tbm_df = db_select("tbm_logs")
    st.dataframe(tbm_df, use_container_width=True)

def page_payroll():
    st.subheader("💰 급여 관리")
    st.write("월별 급여 계산 및 명세서 발행 기능입니다.")

def page_system():
    st.subheader("⚙️ 시스템 관리")
    st.write("시스템 설정 및 메뉴 순서 변경 화면입니다.")

PAGE_FUNCS = {
    "전사현황": page_dashboard,
    "출근입력": page_clockin,
    "근무계획": page_schedule,
    "이력조회": page_history,
    "부서관리": page_departments,
    "직원관리": page_employees,
    "TBM안전관리": page_tbm,
    "급여관리": page_payroll,
    "시스템관리": page_system,
}

def main():
    st.markdown(CUSTOM_CSS, unsafe_allow_html=True)
    with st.sidebar:
        st.markdown(
            f'<div class="brand-head"><div class="brand-emoji">{LOGO_EMOJI}</div>'
            f'<div><div class="brand-name">{APP_TITLE}</div><div class="brand-sub">{APP_SUBTITLE}</div></div></div>',
            unsafe_allow_html=True,
        )
        menu_labels = [m[0] for m in DEFAULT_MENU]
        menu_icons = [m[1] for m in DEFAULT_MENU]
        
        if _HAS_OPTION_MENU:
            selected = option_menu(
                menu_title=None,
                options=menu_labels,
                icons=["bar-chart-fill", "clipboard-check", "calendar-week", "search", "building", "person", "shield-check", "cash-coin", "gear"],
                default_index=0,
                styles={
                    "container": {"padding": "0!important", "background-color": "transparent"},
                    "icon": {"color": "#9ec5fe", "font-size": "14px"},
                    "nav-link": {"font-size": "14px", "text-align": "left", "margin": "2px 0", "color": "#e9eef5"},
                    "nav-link-selected": {"background-color": "#0d6efd", "color": "#ffffff", "font-weight": "700"},
                }
            )
        else:
            selected = st.radio("메뉴 이동", menu_labels)

    func = PAGE_FUNCS.get(selected, page_dashboard)
    func()

if __name__ == "__main__":
    main()
