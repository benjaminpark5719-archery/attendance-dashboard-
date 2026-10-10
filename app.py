# -*- coding: utf-8 -*-
"""
================================================================================
 보물섬수산 HR 시스템 — 통합 app.py (실 Supabase 연동판 / v3.0)
================================================================================
 employees:  emp_no, name, dept_name, scheduled_time, position,
             employment_status, hire_date, base_salary, allowance,
             pay_type, health_cert_expiry, phone ...
 attendance: emp_no, name, dept_name, work_date, actual_start, status
 tbm_records / tbm_confirmations
 부서(departments) 전용 테이블 없음 → employees.dept_name 에서 자동 추출
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

st.set_page_config(page_title="보물섬수산 HR", page_icon="🐟", layout="wide")

APP_TITLE = "보물섬수산"
APP_SUBTITLE = "HR 통합 관리 시스템"
LOGO_EMOJI = "🐟"

LOGIN_ID = "admin"
LOGIN_PW = "admin1234!"
AUTH_TOKEN = "bms-hr-ok"

TB = {
    "employees": "employees",
    "attendance": "attendance",
    "tbm_records": "tbm_records",
    "tbm_confirmations": "tbm_confirmations",
}

COL = {
    "dept": "dept_name",
    "status": "employment_status",
    "scheduled": "scheduled_time",
    "salary": "base_salary",
    "allowance": "allowance",
    "cert": "health_cert_expiry",
}

GRACE_MINUTES = 0
STATUS_COLOR = {"정상": "#198754", "지각": "#dc3545", "결근": "#6c757d", "미출근": "#6c757d"}

DEFAULT_MENU = [
    ("전사현황", "📊"), ("출근입력", "📋"), ("근무계획", "🗓️"), ("이력조회", "🔍"),
    ("부서관리", "🏢"), ("직원관리", "👤"), ("TBM안전관리", "🦺"),
    ("급여관리", "💰"), ("시스템관리", "⚙️"),
]

CUSTOM_CSS = """
<style>
section[data-testid="stSidebar"] { background:#0b1f33; }
section[data-testid="stSidebar"] * { color:#e9eef5; }
.brand-head { display:flex; align-items:center; gap:.55rem; padding:.4rem .2rem 1rem .2rem;
  border-bottom:1px solid rgba(255,255,255,.12); margin-bottom:.8rem; }
.brand-emoji { font-size:1.9rem; }
.brand-name  { font-size:1.15rem; font-weight:800; line-height:1.1; }
.brand-sub   { font-size:.72rem; opacity:.75; }
.metric-card { background:#fff; border:1px solid #eef0f3; border-radius:16px;
  padding:.9rem 1rem; box-shadow:0 2px 10px rgba(16,24,40,.05); }
.metric-card .label { font-size:.8rem; color:#667085; margin-bottom:.15rem; }
.metric-card .value { font-size:1.7rem; font-weight:800; line-height:1; }
.metric-card .delta { font-size:.72rem; color:#98a2b3; }
.badge { display:inline-block; padding:.18rem .55rem; border-radius:999px;
  font-size:.72rem; font-weight:700; color:#fff; }
.callout { border-radius:14px; padding:.8rem 1rem; margin:.3rem 0 1rem 0; font-weight:600; }
.callout.warn { background:#fff4e5; border:1px solid #ffd9a8; color:#8a4b00; }
.callout.ok   { background:#e7f6ec; border:1px solid #b7e2c5; color:#1b6e3a; }
.block-container { padding-top:3.5rem !important; }
.main h1,.main h2,.main h3,div[data-testid="stHeading"] { line-height:1.4; margin-top:.2rem; }
div[data-testid="stDataFrame"] { border-radius:12px; overflow:hidden; }
@media (max-width:640px){ .metric-card .value{font-size:1.35rem;}
  .block-container{padding-left:.6rem;padding-right:.6rem;} }
</style>
"""

def badge(text, color):
    return f'<span class="badge" style="background:{color}">{text}</span>'

def metric_card(label, value, delta=""):
    st.markdown(
        f'<div class="metric-card"><div class="label">{label}</div>'
        f'<div class="value">{value}</div><div class="delta">{delta}</div></div>',
        unsafe_allow_html=True)

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


def db_select(table):
    if not USE_SUPABASE:
        return pd.DataFrame()
    try:
        res = SB.table(TB[table]).select("*").execute()
        return pd.DataFrame(res.data or [])
    except Exception as e:
        st.session_state.setdefault("_errs", []).append(f"{table} 조회: {e}")
        return pd.DataFrame()


def db_insert(table, row: dict):
    if not USE_SUPABASE:
        st.error("DB가 연결되지 않았습니다. 시스템관리에서 연결 상태를 확인하세요.")
        return False
    try:
        SB.table(TB[table]).insert(row).execute()
        return True
    except Exception as e:
        st.error(f"저장 실패: {e}")
        return False


def db_update(table, match: dict, patch: dict):
    if not USE_SUPABASE:
        return False
    try:
        q = SB.table(TB[table]).update(patch)
        for k, v in match.items():
            q = q.eq(k, v)
        q.execute()
        return True
    except Exception as e:
        st.error(f"수정 실패: {e}")
        return False


def db_delete(table, match: dict):
    if not USE_SUPABASE:
        return False
    try:
        q = SB.table(TB[table]).delete()
        for k, v in match.items():
            q = q.eq(k, v)
        q.execute()
        return True
    except Exception as e:
        st.error(f"삭제 실패: {e}")
        return False


def db_upsert_attendance(row: dict):
    if not USE_SUPABASE:
        st.error("DB 미연결.")
        return False
    try:
        SB.table(TB["attendance"]).delete()\
            .eq("emp_no", row["emp_no"]).eq("work_date", row["work_date"]).execute()
        SB.table(TB["attendance"]).insert(row).execute()
        return True
    except Exception as e:
        st.error(f"출근 저장 실패: {e}")
        return False


def compress_image(uploaded_file, max_side=1024, quality=65):
    img = Image.open(uploaded_file)
    if img.mode in ("RGBA", "P", "LA"):
        img = img.convert("RGB")
    img.thumbnail((max_side, max_side))
    buf = io.BytesIO()
    img.save(buf, format="JPEG", quality=quality, optimize=True)
    return buf.getvalue()


def _parse_time(v):
    if v in (None, "", "None") or (isinstance(v, float) and pd.isna(v)):
        return None
    if isinstance(v, time):
        return v
    for fmt in ("%H:%M:%S", "%H:%M"):
        try:
            return datetime.strptime(str(v)[:8] if len(str(v)) > 5 else str(v), fmt).time()
        except ValueError:
            continue
    return None


def judge_attendance(scheduled, actual, grace=GRACE_MINUTES):
    scheduled = _parse_time(scheduled)
    actual = _parse_time(actual)
    if scheduled is None:
        return "정상" if actual else "결근"
    if actual is None:
        return "결근"
    limit = (datetime.combine(date.today(), scheduled) + timedelta(minutes=grace)).time()
    return "지각" if actual > limit else "정상"


def active_employees(emp_df):
    if emp_df.empty or COL["status"] not in emp_df.columns:
        return emp_df
    return emp_df[emp_df[COL["status"]] == "재직"]


def dept_list_from_employees(emp_df):
    if emp_df.empty or COL["dept"] not in emp_df.columns:
        return []
    return sorted([d for d in emp_df[COL["dept"]].dropna().unique().tolist() if str(d).strip()])


def build_shift_template_xlsx(emp_df):
    if not _HAS_OPENPYXL:
        return None
    wb = Workbook(); ws = wb.active; ws.title = "근무계획"
    headers = ["부서", "이름", "월", "화", "수", "목", "금", "토", "일"]
    ws.append(headers)
    fill = PatternFill("solid", fgColor="0D6EFD")
    for c in range(1, len(headers) + 1):
        cell = ws.cell(row=1, column=c)
        cell.font = Font(bold=True, color="FFFFFF"); cell.fill = fill
        cell.alignment = Alignment(horizontal="center")
    for _, e in emp_df.iterrows():
        ws.append([e.get(COL["dept"], ""), e.get("name", "")] + ["휴무"] * 7)
    dv = DataValidation(type="list", formula1='"오전,오후,야간,휴무"', allow_blank=True)
    ws.add_data_validation(dv); dv.add(f"C2:I{max(ws.max_row, 200)}")
    for col, w in zip("ABCDEFGHI", [14, 12, 7, 7, 7, 7, 7, 7, 7]):
        ws.column_dimensions[col].width = w
    buf = io.BytesIO(); wb.save(buf); return buf.getvalue()


def login_gate():
    if not st.session_state.get("auth") and st.query_params.get("auth") == AUTH_TOKEN:
        st.session_state.auth = True
    if st.session_state.get("auth"):
        return True
    st.markdown(CUSTOM_CSS, unsafe_allow_html=True)
    st.markdown(f"## {LOGO_EMOJI} {APP_TITLE} {APP_SUBTITLE}")
    with st.form("login"):
        uid = st.text_input("아이디")
        pw = st.text_input("비밀번호", type="password")
        ok = st.form_submit_button("로그인", use_container_width=True)
    if ok:
        if uid == LOGIN_ID and pw == LOGIN_PW:
            st.session_state.auth = True
            st.query_params["auth"] = AUTH_TOKEN
            st.rerun()
        else:
            st.error("아이디 또는 비밀번호가 올바르지 않습니다.")
    return False


def page_dashboard():
    st.subheader("📊 전사 출근 현황 (실시간)")
    emp = db_select("employees")
    att = db_select("attendance")
    today = str(date.today())

    act = active_employees(emp)
    total = len(act)
    today_att = att[att["work_date"] == today] if (not att.empty and "work_date" in att.columns) else pd.DataFrame()

    normal = late = 0
    if not today_att.empty and "status" in today_att.columns:
        normal = int((today_att["status"] == "정상").sum())
        late = int((today_att["status"] == "지각").sum())
    done = today_att["emp_no"].nunique() if (not today_att.empty and "emp_no" in today_att.columns) else 0
    absent = max(0, total - done)

    c1, c2, c3, c4 = st.columns(4)
    with c1: metric_card("전체 재직", f"{total}명")
    with c2: metric_card("정상 출근", f"{normal}명", f"{normal/total*100:.0f}%" if total else "-")
    with c3: metric_card("지각", f"{late}명")
    with c4: metric_card("미출근/결근", f"{absent}명")

    st.write("")
    if late or absent:
        names = today_att[today_att["status"] == "지각"]["name"].tolist() if not today_att.empty else []
        msg = f"⚠️ 지각 {late}명" + (f" ({', '.join(names)})" if names else "") + f"  ·  미출근/결근 {absent}명 — 확인 필요"
        st.markdown(f'<div class="callout warn">{msg}</div>', unsafe_allow_html=True)
    elif total:
        st.markdown('<div class="callout ok">✅ 금일 지각·결근 없음. 전원 정상입니다.</div>', unsafe_allow_html=True)

    st.markdown("##### 금일 상세 출근 현황")
    if not today_att.empty:
        ren = {"emp_no": "사번", "name": "이름", COL["dept"]: "부서",
               "actual_start": "출근시간", "status": "상태"}
        cols = [c for c in ren if c in today_att.columns]
        st.dataframe(today_att[cols].rename(columns=ren), use_container_width=True, hide_index=True)
    else:
        st.info("오늘 등록된 출근 내역이 없습니다.")


def page_clockin():
    st.subheader("📋 출근입력")
    emp = active_employees(db_select("employees"))
    if emp.empty:
        st.info("등록된 재직 직원이 없습니다.")
        return
    depts = dept_list_from_employees(emp)
    counts = emp[COL["dept"]].value_counts().to_dict()
    labels = [f"{d} ({counts.get(d,0)}명)" for d in depts]
    sel = st.selectbox("부서 선택", labels)
    dept = depts[labels.index(sel)] if sel else None
    members = emp[emp[COL["dept"]] == dept]
    target_date = st.date_input("출근일", value=date.today())

    a1, a2 = st.columns(2)
    bulk = a1.button("⚡ 전체 예정시간대로 일괄 출근", use_container_width=True)
    save = a2.button("💾 출근 데이터 저장", type="primary", use_container_width=True)

    st.divider()
    inputs = {}
    for _, e in members.iterrows():
        sched = _parse_time(e.get(COL["scheduled"]))
        default = sched or time(9, 0)
        c1, c2, c3 = st.columns([2, 2, 2])
        c1.markdown(f"**{e['name']}** · {e['emp_no']}")
        c1.caption(f"예정 {sched.strftime('%H:%M') if sched else '미지정'}")
        if bulk and sched:
            st.session_state[f"ci_{e['emp_no']}"] = sched
        actual = c2.time_input("출근", key=f"ci_{e['emp_no']}", value=default,
                               label_visibility="collapsed")
        status = judge_attendance(sched, actual)
        c3.markdown(badge(status, STATUS_COLOR[status]), unsafe_allow_html=True)
        inputs[e["emp_no"]] = (e, sched, actual, status)

    if save and inputs:
        n = 0
        for emp_no, (e, sched, actual, status) in inputs.items():
            row = {"emp_no": emp_no, "name": e["name"], COL["dept"]: e.get(COL["dept"], ""),
                   "work_date": str(target_date),
                   "actual_start": actual.strftime("%H:%M") if actual else None,
                   "status": status}
            if db_upsert_attendance(row):
                n += 1
        if n:
            st.success(f"{n}건 저장 완료!")
            st.rerun()


def page_schedule():
    st.subheader("🗓️ 근무계획 (순환 탄력근무제)")
    emp = db_select("employees")
    tmpl = build_shift_template_xlsx(emp) if not emp.empty else None
    if tmpl:
        st.download_button("📥 주간 순환근무 표준 엑셀 양식 다운로드", data=tmpl,
                           file_name="보물섬수산_주간근무계획_양식.xlsx",
                           mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                           use_container_width=True)
    st.info("현재 버전은 엑셀 양식 배포까지 지원합니다. 근무조 DB 저장·시각 매트릭스·"
            "출근입력 자동연동은 다음 단계에서 전용 테이블(shift_plans)을 만들어 붙입니다.")
    st.caption("※ 지금도 직원별 scheduled_time(예정 출근시간)을 기준으로 출근입력에서 지각 자동판정이 작동합니다.")


def page_history():
    st.subheader("🔍 출퇴근 이력조회")
    att = db_select("attendance")
    if att.empty:
        st.info("조회할 출퇴근 기록이 없습니다.")
        return
    c1, c2, c3 = st.columns([2,
