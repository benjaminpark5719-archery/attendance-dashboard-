"""
=================================================================
보물섬수산 출근현황 관리 HR 시스템 v2.0
Supabase REST API (HTTPS) 방식 — 포트 연결 문제 원천 차단
=================================================================
"""
import streamlit as st
import pandas as pd
import requests
from datetime import datetime, date
import io
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side

# ─── 페이지 설정 ─────────────────────────────────────────────
st.set_page_config(page_title="보물섬수산 HR", page_icon="🐟", layout="wide")

# ─── Supabase REST API 설정 ──────────────────────────────────
SB_URL = st.secrets["supabase"]["url"]
SB_KEY = st.secrets["supabase"]["key"]
HEADERS = {
    "apikey": SB_KEY,
    "Authorization": f"Bearer {SB_KEY}",
    "Content-Type": "application/json",
}

# ─── DB 유틸 함수 ────────────────────────────────────────────

def sb_select(table, params="", order=""):
    """테이블 조회"""
    url = f"{SB_URL}/rest/v1/{table}?select=*"
    if params:
        url += f"&{params}"
    if order:
        url += f"&order={order}"
    r = requests.get(url, headers=HEADERS)
    return r.json() if r.ok else []

def sb_insert(table, data):
    """단건/다건 삽입"""
    r = requests.post(
        f"{SB_URL}/rest/v1/{table}",
        json=data, headers={**HEADERS, "Prefer": "return=representation"}
    )
    return r.ok, r.text

def sb_upsert(table, data, on_conflict):
    """UPSERT (있으면 업데이트, 없으면 삽입)"""
    r = requests.post(
        f"{SB_URL}/rest/v1/{table}?on_conflict={on_conflict}",
        json=data,
        headers={**HEADERS, "Prefer": "return=representation,resolution=merge-duplicates"}
    )
    return r.ok, r.text

def sb_update(table, data, filters):
    """조건부 업데이트"""
    r = requests.patch(
        f"{SB_URL}/rest/v1/{table}?{filters}",
        json=data, headers={**HEADERS, "Prefer": "return=representation"}
    )
    return r.ok, r.text

def sb_delete(table, filters):
    """조건부 삭제"""
    r = requests.delete(
        f"{SB_URL}/rest/v1/{table}?{filters}", headers=HEADERS
    )
    return r.ok

def sb_health_check():
    """DB 연결 상태 확인"""
    try:
        r = requests.get(
            f"{SB_URL}/rest/v1/departments?select=count",
            headers={**HEADERS, "Prefer": "count=exact"},
            timeout=5
        )
        return r.ok
    except Exception:
        return False

# ─── 초기 데이터 세팅 ────────────────────────────────────────

DEFAULT_DEPTS = ["회계부", "회주방", "식당 홀", "식당 주방",
                 "배송팀", "물류팀", "해썹가공공장", "마트 입점팀"]

def init_departments():
    """기본 부서 데이터 삽입 (중복 무시)"""
    existing = sb_select("departments")
    existing_names = {d["name"] for d in existing}
    new_depts = [{"name": n} for n in DEFAULT_DEPTS if n not in existing_names]
    if new_depts:
        sb_insert("departments", new_depts)

# ─── 로그인 ──────────────────────────────────────────────────

STATUS_OPTIONS = ["출근완료", "지각", "결근", "휴무", "조퇴"]

if 'logged_in' not in st.session_state:
    st.session_state['logged_in'] = False

def show_login():
    st.markdown("<div style='text-align:center; margin-top:3rem;'>"
                "<h1>🐟 보물섬수산</h1>"
                "<h3 style='color:#64748b;'>출근현황 관리 HR 시스템</h3>"
                "</div>", unsafe_allow_html=True)
    col1, col2, col3 = st.columns([1, 1.2, 1])
    with col2:
        st.markdown("---")
        with st.form("login_form"):
            username = st.text_input("아이디", value="admin")
            password = st.text_input("비밀번호", type="password")
            submitted = st.form_submit_button("로그인", use_container_width=True, type="primary")
        if submitted:
            if username == "admin" and password == "admin1234!":
                st.session_state['logged_in'] = True
                st.rerun()
            else:
                st.error("아이디 또는 비밀번호가 올바르지 않습니다.")
        st.caption("초기 계정: admin / admin1234!")

# ─── 엑셀 다운로드 헬퍼 ──────────────────────────────────────

def to_excel(df, sheet_name='데이터'):
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df.to_excel(writer, sheet_name=sheet_name, index=False)
        ws = writer.sheets[sheet_name]
        hfill = PatternFill(start_color="1a365d", end_color="1a365d", fill_type="solid")
        hfont = Font(bold=True, color="FFFFFF", size=11)
        thin = Border(left=Side(style='thin'), right=Side(style='thin'),
                      top=Side(style='thin'), bottom=Side(style='thin'))
        for cell in ws[1]:
            cell.fill = hfill
            cell.font = hfont
            cell.alignment = Alignment(horizontal="center")
        for row in ws.iter_rows(min_row=2, max_row=len(df)+1):
            for cell in row:
                cell.border = thin
                cell.alignment = Alignment(horizontal="center")
        for col in ws.columns:
            max_len = max((len(str(c.value)) for c in col if c.value), default=0)
            ws.column_dimensions[col[0].column_letter].width = min(max_len + 4, 25)
    output.seek(0)
    return output.getvalue()

# ═══════════════════════════════════════════════════════════════
#  메뉴 1: 📊 전사 출근 현황
# ═══════════════════════════════════════════════════════════════

def page_dashboard():
    st.title("📊 전사 출근 현황")
    target_date = st.date_input("기준일", value=date.today())
    target_str = target_date.strftime("%Y-%m-%d")

    employees = sb_select("employees", order="dept_name,emp_no")
    attendance = sb_select("attendance", f"date=eq.{target_str}")

    if not employees:
        st.info("등록된 직원이 없습니다. '직원관리' 메뉴에서 엑셀 업로드를 먼저 해주세요.")
        return

    att_map = {a["emp_no"]: a for a in attendance}
    depts = sb_select("departments", order="name")
    dept_names = [d["name"] for d in depts]

    # 부서별 통계
    stats = []
    total = {"부서": "합계", "총인원": 0, "출근완료": 0, "지각": 0, "결근": 0, "휴무": 0, "미입력": 0}
    for dept in dept_names:
        dept_emps = [e for e in employees if e["dept_name"] == dept]
        if not dept_emps:
            continue
        cnt = len(dept_emps)
        completed = sum(1 for e in dept_emps if att_map.get(e["emp_no"], {}).get("status") == "출근완료")
        late = sum(1 for e in dept_emps if att_map.get(e["emp_no"], {}).get("status") == "지각")
        absent = sum(1 for e in dept_emps if att_map.get(e["emp_no"], {}).get("status") == "결근")
        off = sum(1 for e in dept_emps if att_map.get(e["emp_no"], {}).get("status") == "휴무")
        no_input = cnt - completed - late - absent - off
        working = cnt - off
        rate = f"{round(completed / working * 100, 1)}%" if working > 0 else "0%"
        row = {"부서": dept, "총인원": cnt, "출근완료": completed, "지각": late,
               "결근": absent, "휴무": off, "미입력": no_input, "출근율": rate}
        stats.append(row)
        for k in total:
            if k not in ("부서", "출근율"):
                total[k] = total.get(k, 0) + row.get(k, 0)

    working_total = total["총인원"] - total["휴무"]
    total["출근율"] = f"{round(total['출근완료'] / working_total * 100, 1)}%" if working_total > 0 else "0%"
    stats.append(total)

    # KPI
    c1, c2, c3, c4, c5 = st.columns(5)
    c1.metric("출근율", total.get("출근율", "0%"))
    c2.metric("출근완료", f"{total.get('출근완료', 0)}명")
    c3.metric("지각", f"{total.get('지각', 0)}명")
    c4.metric("결근", f"{total.get('결근', 0)}명")
    c5.metric("미입력", f"{total.get('미입력', 0)}명")

    st.markdown("---")
    st.subheader("부서별 상세")
    st.dataframe(pd.DataFrame(stats), use_container_width=True, hide_index=True)

    # 전체 직원 리스트
    st.markdown("---")
    st.subheader("전체 직원 상세")
    detail = []
    for e in employees:
        a = att_map.get(e["emp_no"], {})
        detail.append({
            "부서": e["dept_name"], "사번": e["emp_no"], "이름": e["name"],
            "출근예정": e.get("scheduled_time", "09:00"),
            "실제출근": a.get("actual_time", ""),
            "상태": a.get("status", "미입력"),
        })
    detail_df = pd.DataFrame(detail)
    st.dataframe(detail_df, use_container_width=True, hide_index=True, height=400)

    st.download_button(
        "📥 엑셀 다운로드", to_excel(detail_df, "출근현황"),
        f"출근현황_{target_str}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

# ═══════════════════════════════════════════════════════════════
#  메뉴 2: 📋 출근 입력
# ═══════════════════════════════════════════════════════════════

def page_attendance_input():
    st.title("📋 부서별 출근 현황 입력")
    target_date = st.date_input("날짜", value=date.today())
    target_str = target_date.strftime("%Y-%m-%d")

    depts = sb_select("departments", order="name")
    if not depts:
        st.warning("등록된 부서가 없습니다.")
        return

    dept_names = [d["name"] for d in depts]
    selected_dept = st.selectbox("🏢 부서 선택", dept_names)

    employees = sb_select("employees", f"dept_name=eq.{selected_dept}", order="emp_no")
    if not employees:
        st.info(f"{selected_dept}에 등록된 직원이 없습니다.")
        return

    attendance = sb_select("attendance", f"date=eq.{target_str}")
    att_map = {a["emp_no"]: a for a in attendance}

    st.markdown(f"**{selected_dept}** — {len(employees)}명")
    st.markdown("---")

    # 헤더
    cols = st.columns([1.5, 1.5, 1.2, 1.2, 1.5])
    for col, h in zip(cols, ["이름", "사번", "출근예정", "실제출근", "근무상태"]):
        col.markdown(f"**{h}**")
    st.divider()

    new_data = {}
    for emp in employees:
        existing = att_map.get(emp["emp_no"], {})
        cols = st.columns([1.5, 1.5, 1.2, 1.2, 1.5])
        with cols[0]:
            st.write(f"👤 {emp['name']}")
        with cols[1]:
            st.write(emp["emp_no"])
        with cols[2]:
            st.write(emp.get("scheduled_time", "09:00"))
        with cols[3]:
            actual = st.text_input(
                "시간", value=existing.get("actual_time", ""),
                placeholder="HH:MM", key=f"at_{emp['emp_no']}",
                label_visibility="collapsed"
            )
        with cols[4]:
            current_status = existing.get("status", "")
            idx = (STATUS_OPTIONS.index(current_status) + 1) if current_status in STATUS_OPTIONS else 0
            status = st.selectbox(
                "상태", [""] + STATUS_OPTIONS, index=idx,
                key=f"st_{emp['emp_no']}", label_visibility="collapsed"
            )
        new_data[emp["emp_no"]] = {"actual_time": actual, "status": status}

    st.markdown("---")
    c1, c2, c3 = st.columns([1, 1, 1])
    with c2:
        if st.button("💾 저장", use_container_width=True, type="primary"):
            records = []
            for emp_no, vals in new_data.items():
                if vals["status"]:  # 상태가 선택된 것만
                    records.append({
                        "date": target_str,
                        "emp_no": emp_no,
                        "actual_time": vals["actual_time"],
                        "status": vals["status"]
                    })
            if records:
                ok, msg = sb_upsert("attendance", records, "date,emp_no")
                if ok:
                    st.success(f"✅ {len(records)}명의 출근 기록이 저장되었습니다!")
                    st.rerun()
                else:
                    st.error(f"저장 실패: {msg}")
            else:
                st.warning("저장할 데이터가 없습니다. 근무상태를 선택해 주세요.")

# ═══════════════════════════════════════════════════════════════
#  메뉴 3: 🔍 이력 조회
# ═══════════════════════════════════════════════════════════════

def page_history():
    st.title("🔍 출근 이력 조회")

    c1, c2, c3 = st.columns(3)
    with c1:
        date_from = st.date_input("시작일", value=date.today().replace(day=1))
    with c2:
        date_to = st.date_input("종료일", value=date.today())
    with c3:
        depts = sb_select("departments", order="name")
        dept_options = ["전체"] + [d["name"] for d in depts]
        sel_dept = st.selectbox("부서", dept_options)

    filters = f"date=gte.{date_from}&date=lte.{date_to}"
    logs = sb_select("attendance", filters, order="date.desc,emp_no")

    if not logs:
        st.info("해당 기간의 출근 이력이 없습니다.")
        return

    employees = sb_select("employees")
    emp_map = {e["emp_no"]: e for e in employees}

    rows = []
    for log in logs:
        emp = emp_map.get(log["emp_no"], {})
        dept = emp.get("dept_name", "-")
        if sel_dept != "전체" and dept != sel_dept:
            continue
        rows.append({
            "날짜": log["date"], "부서": dept,
            "사번": log["emp_no"], "이름": emp.get("name", "-"),
            "출근예정": emp.get("scheduled_time", ""),
            "실제출근": log.get("actual_time", ""),
            "상태": log.get("status", ""),
        })

    if rows:
        df = pd.DataFrame(rows)
        st.dataframe(df, use_container_width=True, hide_index=True, height=500)
        st.caption(f"총 {len(rows)}건")
        st.download_button(
            "📥 엑셀 다운로드", to_excel(df, "출근이력"),
            f"출근이력_{date_from}_{date_to}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
    else:
        st.info("필터 조건에 맞는 이력이 없습니다.")

# ═══════════════════════════════════════════════════════════════
#  메뉴 4: 🏢 부서 관리
# ═══════════════════════════════════════════════════════════════

def page_departments():
    st.title("🏢 부서 관리")

    tab1, tab2 = st.tabs(["📋 부서 목록", "➕ 부서 등록"])

    with tab1:
        depts = sb_select("departments", order="name")
        if depts:
            for dept in depts:
                employees = sb_select("employees", f"dept_name=eq.{dept['name']}")
                emp_count = len(employees)
                c1, c2, c3 = st.columns([3, 1, 1])
                c1.write(f"🏢 **{dept['name']}** — 직원 {emp_count}명")
                with c3:
                    if emp_count == 0:
                        if st.button("🗑️ 삭제", key=f"del_{dept['id']}"):
                            sb_delete("departments", f"id=eq.{dept['id']}")
                            st.success(f"{dept['name']} 삭제 완료")
                            st.rerun()
                    else:
                        st.caption("삭제불가")
        else:
            st.info("등록된 부서가 없습니다.")

    with tab2:
        with st.form("new_dept"):
            new_name = st.text_input("새 부서명")
            if st.form_submit_button("➕ 등록", type="primary"):
                if new_name:
                    ok, msg = sb_insert("departments", {"name": new_name.strip()})
                    if ok:
                        st.success(f"'{new_name}' 부서가 등록되었습니다.")
                        st.rerun()
                    else:
                        st.error(f"등록 실패 (중복?): {msg}")
                else:
                    st.warning("부서명을 입력하세요.")

# ═══════════════════════════════════════════════════════════════
#  메뉴 5: 👤 직원 관리 (엑셀 업로드)
# ═══════════════════════════════════════════════════════════════

def page_employees():
    st.title("👤 직원 관리")

    tab1, tab2, tab3 = st.tabs(["📥 엑셀 업로드", "📋 직원 목록", "🗑️ 직원 삭제"])

    with tab1:
        st.subheader("엑셀/CSV 파일로 직원 일괄 등록")
        st.markdown("""
엑셀(또는 CSV) 파일에 아래 4개 컬럼이 있어야 합니다:

| 사번 | 이름 | 부서 | 출근예정시간 |
|------|------|------|------------|
| ACC-001 | 홍길동 | 회계부 | 09:00 |
| DLV-001 | 김철수 | 배송팀 | 08:00 |
        """)

        uploaded = st.file_uploader("파일 선택", type=["xlsx", "csv"])
        if uploaded:
            try:
                if uploaded.name.endswith('.csv'):
                    df = pd.read_csv(uploaded)
                else:
                    df = pd.read_excel(uploaded)

                required_cols = {"사번", "이름", "부서", "출근예정시간"}
                if not required_cols.issubset(set(df.columns)):
                    st.error(f"필수 컬럼이 없습니다: {required_cols - set(df.columns)}")
                else:
                    st.write(f"📊 미리보기 — 총 {len(df)}명:")
                    st.dataframe(df.head(10), use_container_width=True, hide_index=True)

                    if st.button("🚀 데이터베이스에 저장하기", type="primary"):
                        records = []
                        for _, row in df.iterrows():
                            records.append({
                                "emp_no": str(row["사번"]).strip(),
                                "name": str(row["이름"]).strip(),
                                "dept_name": str(row["부서"]).strip(),
                                "scheduled_time": str(row["출근예정시간"]).strip()
                            })
                        ok, msg = sb_upsert("employees", records, "emp_no")
                        if ok:
                            st.success(f"🎉 {len(records)}명이 저장되었습니다!")
                            st.rerun()
                        else:
                            st.error(f"저장 실패: {msg}")
            except Exception as e:
                st.error(f"파일 처리 에러: {e}")

    with tab2:
        st.subheader("등록된 직원 목록")
        c1, _ = st.columns([1, 3])
        with c1:
            depts = sb_select("departments", order="name")
            dept_filter = st.selectbox("부서 필터", ["전체"] + [d["name"] for d in depts], key="emp_filter")

        if dept_filter == "전체":
            employees = sb_select("employees", order="dept_name,emp_no")
        else:
            employees = sb_select("employees", f"dept_name=eq.{dept_filter}", order="emp_no")

        if employees:
            df = pd.DataFrame(employees)
            display_df = df[["emp_no", "name", "dept_name", "scheduled_time"]].rename(
                columns={"emp_no": "사번", "name": "이름", "dept_name": "부서", "scheduled_time": "출근예정시간"}
            )
            st.dataframe(display_df, use_container_width=True, hide_index=True)
            st.caption(f"총 {len(employees)}명")
        else:
            st.info("등록된 직원이 없습니다.")

    with tab3:
        st.subheader("직원 삭제")
        emp_no_del = st.text_input("삭제할 사번 입력", placeholder="예: ACC-001")
        if st.button("🗑️ 삭제 실행", type="secondary"):
            if emp_no_del:
                ok = sb_delete("employees", f"emp_no=eq.{emp_no_del.strip()}")
                if ok:
                    st.success(f"사번 {emp_no_del} 삭제 완료")
                    st.rerun()
                else:
                    st.error("삭제 실패: 사번을 확인하세요.")

# ═══════════════════════════════════════════════════════════════
#  메뉴 6: 💰 급여 관리
# ═══════════════════════════════════════════════════════════════

def page_salary():
    st.title("💰 급여 관리")
    st.info("급여 관리 기능은 직원 데이터 안정화 후 추가 개발 예정입니다.")

    employees = sb_select("employees", order="dept_name,emp_no")
    if employees:
        df = pd.DataFrame(employees)
        display_df = df[["emp_no", "name", "dept_name"]].rename(
            columns={"emp_no": "사번", "name": "이름", "dept_name": "부서"}
        )
        display_df["기본급"] = ""
        display_df["수당"] = ""
        display_df["공제"] = ""
        display_df["실수령액"] = ""
        st.dataframe(display_df, use_container_width=True, hide_index=True)
        st.caption("급여 데이터 입력 기능은 다음 버전에서 제공됩니다.")

# ═══════════════════════════════════════════════════════════════
#  메뉴 7: ⚙️ 계정 관리
# ═══════════════════════════════════════════════════════════════

def page_settings():
    st.title("⚙️ 계정 관리")

    st.subheader("현재 계정 정보")
    st.write("**아이디:** admin")
    st.write("**역할:** 총괄관리자")

    st.markdown("---")
    st.subheader("🔑 비밀번호 변경")
    st.info("비밀번호 관리 기능은 다음 버전에서 제공됩니다.")

    st.markdown("---")
    st.subheader("📊 시스템 상태")
    if sb_health_check():
        st.success("✅ Supabase DB 연결: 정상")
    else:
        st.error("❌ Supabase DB 연결: 실패")

    emp_count = len(sb_select("employees"))
    dept_count = len(sb_select("departments"))
    att_count = len(sb_select("attendance"))
    c1, c2, c3 = st.columns(3)
    c1.metric("부서 수", dept_count)
    c2.metric("직원 수", emp_count)
    c3.metric("출근 기록 수", att_count)

# ═══════════════════════════════════════════════════════════════
#  메인 라우터
# ═══════════════════════════════════════════════════════════════

def main():
    if not st.session_state['logged_in']:
        # DB 연결 체크
        if not sb_health_check():
            st.error("❌ Supabase DB 연결 실패. Secrets 설정을 확인하세요.")
        show_login()
        return

    # 초기 부서 데이터 확인
    init_departments()

    # 사이드바 메뉴
    st.sidebar.markdown("### 🐟 보물섬수산 HR")
    st.sidebar.markdown("---")
    menu = st.sidebar.radio(
        "메뉴",
        ["📊 전사현황", "📋 출근입력", "🔍 이력조회",
         "🏢 부서관리", "👤 직원관리", "💰 급여관리", "⚙️ 계정관리"]
    )
    st.sidebar.markdown("---")
    if st.sidebar.button("🔓 로그아웃", use_container_width=True):
        st.session_state['logged_in'] = False
        st.rerun()
    st.sidebar.caption("v2.0 | Supabase REST API")

    # 페이지 라우팅
    pages = {
        "📊 전사현황": page_dashboard,
        "📋 출근입력": page_attendance_input,
        "🔍 이력조회": page_history,
        "🏢 부서관리": page_departments,
        "👤 직원관리": page_employees,
        "💰 급여관리": page_salary,
        "⚙️ 계정관리": page_settings,
    }
    pages[menu]()


if __name__ == "__main__":
    main()
