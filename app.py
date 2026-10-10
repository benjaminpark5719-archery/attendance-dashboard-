# -*- coding: utf-8 -*-
"""
================================================================================
 보물섬수산 HR 시스템 — 통합 app.py (실 Supabase 연동판 / v3.0)
================================================================================
 실제 DB 스키마(setup.sql)에 맞춰 작성:
   employees:  emp_no, name, dept_name, scheduled_time, position,
               employment_status, hire_date, base_salary, allowance,
               pay_type, health_cert_expiry, phone ...
   attendance: emp_no, name, dept_name, work_date, actual_start, status ...
   tbm_records:       tbm_date, department, title, content, photo_data
   tbm_confirmations: tbm_id, emp_no, emp_name, confirmed_at
 부서(departments) 전용 테이블이 없으므로 employees.dept_name 에서 자동 추출.

 로그인 유지: URL 쿼리 파라미터
 헤더 잘림 방지: block-container padding-top 3.5rem
 사진 업로드: Pillow 압축 후 base64 (photo_data 컬럼 TEXT)
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

# ──────────────────────────────────────────────────────────────────────────────
# 기본 설정
# ──────────────────────────────────────────────────────────────────────────────
st.set_page_config(page_title="보물섬수산 HR", page_icon="🐟", layout="wide")

APP_TITLE = "보물섬수산"
APP_SUBTITLE = "HR 통합 관리 시스템"
LOGO_EMOJI = "🐟"

LOGIN_ID = "admin"
LOGIN_PW = "admin1234!"
AUTH_TOKEN = "bms-hr-ok"

# 실제 테이블명 (필요 시 여기만 수정)
TB = {
    "employees": "employees",
    "attendance": "attendance",
    "tbm_records": "tbm_records",
    "tbm_confirmations": "tbm_confirmations",
}

# 실제 컬럼명 매핑 (setup.sql 기준) — DB 컬럼이 바뀌면 여기만 고치면 됨
COL = {
    "dept": "dept_name",            # 부서
    "status": "employment_status",  # 재직/퇴직
    "scheduled": "scheduled_time",  # 예정 출근시간
    "salary": "base_salary",        # 기본급
    "allowance": "allowance",       # 수당
    "cert": "health_cert_expiry",   # 보건증 만료
}

GRACE_MINUTES = 0
STATUS_COLOR = {"정상": "#198754", "지각": "#dc3545", "결근": "#6c757d", "미출근": "#6c757d"}

DEFAULT_MENU = [
    ("전사현황", "📊"), ("출근입력", "📋"), ("근무계획", "🗓️"), ("이력조회", "🔍"),
    ("부서관리", "🏢"), ("직원관리", "👤"), ("TBM안전관리", "🦺"),
    ("급여관리", "💰"), ("시스템관리", "⚙️"),
]

# ──────────────────────────────────────────────────────────────────────────────
# 스타일
# ──────────────────────────────────────────────────────────────────────────────
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
/* 사이드바 메뉴(streamlit-option-menu) 글자 상시 표시 */
section[data-testid="stSidebar"] .nav-link,
section[data-testid="stSidebar"] .nav-link span { color:#e9eef5 !important; }
section[data-testid="stSidebar"] div[role="button"] { background-color:#0b1f33 !important; }
section[data-testid="stSidebar"] ul { background-color:#0b1f33 !important; }
</style>
"""

def badge(text, color):
    return f'<span class="badge" style="background:{color}">{text}</span>'

def metric_card(label, value, delta=""):
    st.markdown(
        f'<div class="metric-card"><div class="label">{label}</div>'
        f'<div class="value">{value}</div><div class="delta">{delta}</div></div>',
        unsafe_allow_html=True)

# ──────────────────────────────────────────────────────────────────────────────
# Supabase 연결
# ──────────────────────────────────────────────────────────────────────────────
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
    """테이블 전체 조회 → DataFrame. 연결 안 되면 빈 DF."""
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
    """(emp_no, work_date) 기준 중복 제거 후 저장."""
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

# ──────────────────────────────────────────────────────────────────────────────
# 유틸
# ──────────────────────────────────────────────────────────────────────────────
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
    """부서 전용 테이블이 없으므로 직원의 dept_name에서 부서 목록 자동 추출."""
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

# ──────────────────────────────────────────────────────────────────────────────
# 인증
# ──────────────────────────────────────────────────────────────────────────────
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

# ──────────────────────────────────────────────────────────────────────────────
# ① 전사현황
# ──────────────────────────────────────────────────────────────────────────────
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

# ──────────────────────────────────────────────────────────────────────────────
# ② 출근입력
# ──────────────────────────────────────────────────────────────────────────────
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

# ──────────────────────────────────────────────────────────────────────────────
# ③ 근무계획
# ──────────────────────────────────────────────────────────────────────────────
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

# ──────────────────────────────────────────────────────────────────────────────
# ④ 이력조회
# ──────────────────────────────────────────────────────────────────────────────
def page_history():
    st.subheader("🔍 출퇴근 이력조회")
    att = db_select("attendance")
    if att.empty:
        st.info("조회할 출퇴근 기록이 없습니다.")
        return
    c1, c2, c3 = st.columns([2, 2, 2])
    kw = c1.text_input("이름 / 사번 검색")
    sf = c2.multiselect("상태", ["정상", "지각", "결근"], default=["정상", "지각", "결근"])
    dr = c3.date_input("기간", value=(date.today() - timedelta(days=30), date.today()))

    df = att.copy()
    if "work_date" in df.columns and isinstance(dr, tuple) and len(dr) == 2:
        df = df[(df["work_date"] >= str(dr[0])) & (df["work_date"] <= str(dr[1]))]
    if kw:
        df = df[df["name"].astype(str).str.contains(kw, na=False) |
                df["emp_no"].astype(str).str.contains(kw, na=False)]
    if sf and "status" in df.columns:
        df = df[df["status"].isin(sf)]

    total = len(df)
    late = int((df["status"] == "지각").sum()) if "status" in df.columns else 0
    absent = int((df["status"] == "결근").sum()) if "status" in df.columns else 0
    rate = (total - absent) / total * 100 if total else 0
    m1, m2, m3 = st.columns(3)
    with m1: metric_card("총 출근율", f"{rate:.1f}%", f"{total}건")
    with m2: metric_card("지각 합계", f"{late}건")
    with m3: metric_card("결근 합계", f"{absent}건")

    ren = {"work_date": "일자", "emp_no": "사번", "name": "이름", COL["dept"]: "부서",
           "actual_start": "출근", "status": "상태"}
    cols = [c for c in ren if c in df.columns]
    view = df[cols].rename(columns=ren)
    if "일자" in view.columns:                      # 있을 때만 정렬 (KeyError 방지)
        view = view.sort_values("일자", ascending=False)
    st.dataframe(view, use_container_width=True, hide_index=True)

# ──────────────────────────────────────────────────────────────────────────────
# ⑤ 부서관리 (dept_name 기반)
# ──────────────────────────────────────────────────────────────────────────────
def page_departments():
    st.subheader("🏢 부서 관리")
    emp = db_select("employees")
    depts = dept_list_from_employees(emp)
    if not depts:
        st.info("직원 데이터에서 부서를 찾을 수 없습니다.")
        return
    counts = emp[COL["dept"]].value_counts().to_dict()
    rows = [{"부서명": d, "인원수": counts.get(d, 0)} for d in depts]
    st.dataframe(pd.DataFrame(rows), use_container_width=True, hide_index=True)
    st.caption("※ 부서는 직원 정보(부서명)에서 자동 집계됩니다. 부서명을 바꾸려면 직원관리에서 "
               "해당 직원들의 부서를 수정하세요. (부서 전용 테이블은 추후 추가 예정)")

    with st.expander("✏️ 부서명 일괄 변경 (그 부서 소속 전원 이동)"):
        old = st.selectbox("기존 부서", depts)
        new = st.text_input("새 부서명")
        if st.button("일괄 변경 실행", type="primary"):
            if new.strip():
                ok = db_update("employees", {COL["dept"]: old}, {COL["dept"]: new.strip()})
                if ok:
                    st.success(f"'{old}' → '{new.strip()}' 로 변경했습니다.")
                    st.rerun()
            else:
                st.warning("새 부서명을 입력하세요.")

# ──────────────────────────────────────────────────────────────────────────────
# ⑥ 직원관리
# ──────────────────────────────────────────────────────────────────────────────
def _health_badge(expiry):
    try:
        d = datetime.strptime(str(expiry)[:10], "%Y-%m-%d").date()
    except Exception:
        return badge("미등록", "#6c757d")
    left = (d - date.today()).days
    if left < 0: return badge(f"만료 {abs(left)}일", "#dc3545")
    if left <= 30: return badge(f"D-{left}", "#fd7e14")
    return badge("정상", "#198754")


def page_employees():
    st.subheader("👤 직원 관리")
    emp = db_select("employees")
    depts = dept_list_from_employees(emp) or ["회계부", "회주방", "식당 홀", "식당 주방",
                                              "배송팀", "물류팀", "해썹가공공장", "마트 입점팀"]

    with st.expander("➕ 신규 직원 등록 / ✏️ 수정"):
        t1, t2 = st.tabs(["신규 등록", "정보 수정/삭제"])
        with t1:
            with st.form("add_emp"):
                c1, c2, c3 = st.columns(3)
                emp_no = c1.text_input("사번")
                name = c2.text_input("이름")
                dept = c3.selectbox("부서", depts)
                c4, c5, c6 = st.columns(3)
                pos = c4.text_input("직급", value="사원")
                sched = c5.text_input("예정 출근시간", value="09:00")
                salary = c6.number_input("기본급(원)", min_value=0, value=0, step=10000)
                c7, c8 = st.columns(2)
                hire = c7.date_input("입사일", value=date.today())
                cert = c8.text_input("보건증 만료일(YYYY-MM-DD)", value="")
                if st.form_submit_button("등록", type="primary"):
                    if not emp_no.strip() or not name.strip():
                        st.error("사번과 이름은 반드시 입력해야 합니다.")
                    else:
                        row = {"emp_no": emp_no.strip(), "name": name.strip(), COL["dept"]: dept,
                               "position": pos, COL["scheduled"]: sched,
                               COL["salary"]: int(salary), "hire_date": str(hire),
                               COL["cert"]: cert, COL["status"]: "재직"}
                        if db_insert("employees", row):
                            st.success(f"'{name}' 등록 완료."); st.rerun()
        with t2:
            if not emp.empty:
                opts = emp["emp_no"].tolist()
                sel = st.selectbox("직원 선택(사번)", opts,
                    format_func=lambda x: f"{emp[emp['emp_no']==x]['name'].values[0]} ({x})")
                cur = emp[emp["emp_no"] == sel].iloc[0]
                with st.form("edit_emp"):
                    c1, c2, c3 = st.columns(3)
                    en = c1.text_input("이름", value=str(cur.get("name", "")))
                    ed = c2.selectbox("부서", depts,
                        index=depts.index(cur.get(COL["dept"])) if cur.get(COL["dept"]) in depts else 0)
                    es = c3.selectbox("상태", ["재직", "퇴직"],
                        index=0 if cur.get(COL["status"]) == "재직" else 1)
                    c4, c5 = st.columns(2)
                    esc = c4.text_input("예정 출근시간", value=str(cur.get(COL["scheduled"], "")))
                    ec = c5.text_input("보건증 만료일", value=str(cur.get(COL["cert"], "")))
                    csave, cdel = st.columns(2)
                    if csave.form_submit_button("수정 저장", type="primary"):
                        patch = {"name": en, COL["dept"]: ed, COL["status"]: es,
                                 COL["scheduled"]: esc, COL["cert"]: ec}
                        if db_update("employees", {"emp_no": sel}, patch):
                            st.success("수정 완료."); st.rerun()
                    if cdel.form_submit_button("🗑️ 삭제"):
                        if db_delete("employees", {"emp_no": sel}):
                            st.warning("삭제 완료."); st.rerun()

    if emp.empty:
        st.info("등록된 직원이 없습니다.")
        return
    kw = st.text_input("🔎 이름 / 사번 / 부서 검색")
    df = emp.copy()
    if kw:
        df = df[df["name"].astype(str).str.contains(kw, na=False) |
                df["emp_no"].astype(str).str.contains(kw, na=False) |
                df[COL["dept"]].astype(str).str.contains(kw, na=False)]

    # 보건증 경고
    if COL["cert"] in df.columns:
        warn = []
        for _, r in df.iterrows():
            try:
                left = (datetime.strptime(str(r[COL["cert"]])[:10], "%Y-%m-%d").date() - date.today()).days
                if left <= 30:
                    warn.append(f"⚠️ {r['name']}({r.get(COL['dept'],'')}) · {r[COL['cert']]} "
                                f"({'만료' if left < 0 else f'D-{left}'})")
            except Exception:
                pass
        if warn:
            st.markdown('<div class="callout warn">🚨 보건증 갱신 필요 (30일 이내):<br>'
                        + "<br>".join(warn) + '</div>', unsafe_allow_html=True)

    ren = {"emp_no": "사번", "name": "이름", COL["dept"]: "부서", "position": "직급",
           COL["scheduled"]: "예정출근", COL["status"]: "상태", COL["cert"]: "보건증만료"}
    cols = [c for c in ren if c in df.columns]
    st.dataframe(df[cols].rename(columns=ren), use_container_width=True, hide_index=True)

# ──────────────────────────────────────────────────────────────────────────────
# ⑦ TBM 안전관리 (tbm_records + tbm_confirmations)
# ──────────────────────────────────────────────────────────────────────────────
def page_tbm():
    st.subheader("🦺 TBM 안전보건일지")
    emp = db_select("employees")
    depts = dept_list_from_employees(emp)

    with st.expander("➕ TBM 일지 작성"):
        with st.form("add_tbm"):
            c1, c2 = st.columns(2)
            tdate = c1.date_input("점검일자", value=date.today())
            dept = c2.selectbox("부서", depts if depts else ["전체"])
            title = st.text_input("TBM 주제 (필수)")
            content = st.text_area("교육 내용 / 위험요인 및 대책")
            photo = st.file_uploader("현장 사진 (자동 압축)", type=["jpg", "jpeg", "png"])
            if st.form_submit_button("일지 저장", type="primary"):
                if not title.strip():
                    st.error("TBM 주제를 입력해야 저장할 수 있습니다.")
                else:
                    photo_b64 = ""
                    if photo is not None:
                        try:
                            photo_b64 = base64.b64encode(compress_image(photo)).decode()
                        except Exception as e:
                            st.warning(f"이미지 처리 실패: {e}")
                    row = {"tbm_date": str(tdate), "department": dept, "title": title.strip(),
                           "content": content, "photo_data": photo_b64}
                    if db_insert("tbm_records", row):
                        st.success("TBM 일지 등록 완료."); st.rerun()

    logs = db_select("tbm_records")
    if logs.empty:
        st.info("등록된 TBM 일지가 없습니다.")
        return
    for _, log in logs.sort_values("tbm_date", ascending=False).iterrows():
        with st.expander(f"[{log['tbm_date']}] {log.get('department','')} · {log.get('title','')}"):
            st.write(log.get("content", ""))
            pd_ = log.get("photo_data", "")
            if isinstance(pd_, str) and pd_:
                try:
                    st.image(base64.b64decode(pd_), width=280)
                except Exception:
                    pass

# ──────────────────────────────────────────────────────────────────────────────
# ⑧ 급여관리
# ──────────────────────────────────────────────────────────────────────────────
def page_payroll():
    st.subheader("💰 급여 관리 (시뮬레이터)")
    emp = db_select("employees")
    att = db_select("attendance")
    if emp.empty:
        st.info("직원 데이터가 없습니다.")
        return
    ym = st.text_input("정산 월 (YYYY-MM)", value=date.today().strftime("%Y-%m"))
    month_att = att[att["work_date"].astype(str).str.startswith(ym)] if (not att.empty and "work_date" in att.columns) else pd.DataFrame()

    rows = []
    for _, e in active_employees(emp).iterrows():
        rec = month_att[month_att["emp_no"] == e["emp_no"]] if not month_att.empty else pd.DataFrame()
        work_days = len(rec[rec["status"].isin(["정상", "지각"])]) if (not rec.empty and "status" in rec.columns) else 0
        base = int(e.get(COL["salary"], 0) or 0)
        allow = int(e.get(COL["allowance"], 0) or 0)
        total = base + allow
        rows.append({"사번": e["emp_no"], "이름": e["name"], "부서": e.get(COL["dept"], ""),
                     "근무일수": f"{work_days}일", "기본급": f"{base:,}원",
                     "수당": f"{allow:,}원", "예상 지급액": f"{total:,}원"})
    st.dataframe(pd.DataFrame(rows), use_container_width=True, hide_index=True)
    st.caption("※ 기본급·수당은 직원 정보(base_salary/allowance) 기준. 근무일수는 해당 월 출근기록에서 집계. "
               "연장·휴일수당 자동계산과 명세서 PDF는 다음 단계에서 추가합니다.")

# ──────────────────────────────────────────────────────────────────────────────
# ⑨ 시스템관리
# ──────────────────────────────────────────────────────────────────────────────
def page_system():
    st.subheader("⚙️ 시스템 관리")
    st.markdown("##### 연결 상태")
    if USE_SUPABASE:
        st.success("Supabase 연결: 🟢 연결됨")
        emp = db_select("employees")
        st.write(f"- 직원 테이블 로드: **{len(emp)}명**")
        st.write(f"- 감지된 부서: **{', '.join(dept_list_from_employees(emp)) or '없음'}**")
    else:
        st.error("Supabase 연결: 🔴 미연결 — Secrets(SUPABASE_URL/SUPABASE_KEY)를 확인하세요.")
    if st.session_state.get("_errs"):
        with st.expander("오류 로그"):
            for e in st.session_state["_errs"][-20:]:
                st.caption(e)

# ──────────────────────────────────────────────────────────────────────────────
# 라우터
# ──────────────────────────────────────────────────────────────────────────────
PAGE_FUNCS = {
    "전사현황": page_dashboard, "출근입력": page_clockin, "근무계획": page_schedule,
    "이력조회": page_history, "부서관리": page_departments, "직원관리": page_employees,
    "TBM안전관리": page_tbm, "급여관리": page_payroll, "시스템관리": page_system,
}


def main():
    st.markdown(CUSTOM_CSS, unsafe_allow_html=True)
    if not login_gate():
        return
    labels = [m[0] for m in DEFAULT_MENU]
    icons = [m[1] for m in DEFAULT_MENU]
    with st.sidebar:
        st.markdown(
            f'<div class="brand-head"><div class="brand-emoji">{LOGO_EMOJI}</div>'
            f'<div><div class="brand-name">{APP_TITLE}</div>'
            f'<div class="brand-sub">{APP_SUBTITLE}</div></div></div>',
            unsafe_allow_html=True)
        if _HAS_OPTION_MENU:
            choice = option_menu(None, [f"{ic} {lb}" for lb, ic in zip(labels, icons)],
                default_index=0,
                styles={"container": {"background-color": "#0b1f33", "padding": "4px 0"},
                        "nav-link": {"color": "#e9eef5", "font-size": "14px",
                                     "font-weight": "600", "background-color": "#0b1f33",
                                     "margin": "2px 6px", "border-radius": "8px"},
                        "nav-link-selected": {"background-color": "#0d6efd", "color": "#ffffff"}})
            selected = labels[[f"{ic} {lb}" for lb, ic in zip(labels, icons)].index(choice)]
        else:
            selected = st.radio("메뉴 이동", labels,
                format_func=lambda l: f"{icons[labels.index(l)]} {l}",
                label_visibility="collapsed")
        st.divider()
        if st.button("로그아웃", use_container_width=True):
            st.session_state.auth = False
            try:
                del st.query_params["auth"]
            except Exception:
                st.query_params.clear()
            st.rerun()
    PAGE_FUNCS.get(selected, page_dashboard)()


if __name__ == "__main__":
    main()
