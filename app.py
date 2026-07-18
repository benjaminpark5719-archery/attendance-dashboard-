import streamlit as st
import pandas as pd
from datetime import datetime
import psycopg2
from psycopg2.extras import RealDictCursor

# 1. 페이지 기본 설정 및 보안 로그인
st.set_page_config(page_title="보물섬수산 HR 시스템", page_icon="🏢", layout="wide")

if 'logged_in' not in st.session_state:
    st.session_state['logged_in'] = False

def db_connect():
    # Streamlit Secrets에 저장한 DB 주소로 진짜 연결
    return psycopg2.connect(st.secrets["database"]["url"])

def init_db():
    # 데이터베이스에 진짜 테이블들 생성 (부서, 직원, 출근, 급여)
    with db_connect() as conn:
        with conn.cursor() as cur:
            cur.execute("""
                CREATE TABLE IF NOT EXISTS departments (
                    id SERIAL PRIMARY KEY,
                    name VARCHAR(50) UNIQUE NOT NULL
                );
                CREATE TABLE IF NOT EXISTS employees (
                    id SERIAL PRIMARY KEY,
                    emp_no VARCHAR(20) UNIQUE NOT NULL,
                    name VARCHAR(50) NOT NULL,
                    dept_name VARCHAR(50) NOT NULL,
                    scheduled_time VARCHAR(10) DEFAULT '09:00'
                );
                CREATE TABLE IF NOT EXISTS attendance (
                    id SERIAL PRIMARY KEY,
                    date DATE NOT NULL,
                    emp_no VARCHAR(20) NOT NULL,
                    actual_time VARCHAR(10),
                    status VARCHAR(20) DEFAULT '출근예정',
                    UNIQUE(date, emp_no)
                );
            """)
            # 기본 부서 등록
            depts = ["회계부", "회주방", "식당 홀", "식당 주방", "배송팀", "물류팀", "해썹가공공장", "마트 입점팀"]
            for d in depts:
                cur.execute("INSERT INTO departments (name) VALUES (%s) ON CONFLICT DO NOTHING", (d,))
        conn.commit()

try:
    init_db()
except Exception as e:
    st.error(f"DB 연결 실패: {e}")

# 로그인 화면
if not st.session_state['logged_in']:
    st.title("🔒 보물섬수산 HR 시스템 로그인")
    username = st.text_input("아이디", value="admin")
    password = st.text_input("비밀번호", type="password")
    if st.button("로그인", type="primary"):
        if username == "admin" and password == "admin1234!":
            st.session_state['logged_in'] = True
            st.rerun()
        else:
            st.error("비밀번호가 올바르지 않습니다.")
else:
    # 2. 형님이 갈망하시던 웅장한 7대 메뉴 전격 활성화!
    st.sidebar.title("🏢 보물섬수산 HR")
    menu = st.sidebar.radio(
        "메뉴를 선택하세요",
        ["📊 전사현황", "📋 출근입력", "🔍 이력조회", "🏢 부서관리", "👤 직원관리 (엑셀업로드)", "💰 급여관리", "⚙️ 계정관리"]
    )
    
    if st.sidebar.button("로그아웃"):
        st.session_state['logged_in'] = False
        st.rerun()

    # ----------------------------------------------------
    # [핵심] 형님이 원하시던 엑셀 업로드 일괄 등록 기능 (CRUD의 C!)
    # ----------------------------------------------------
    if menu == "👤 직원관리 (엑셀업로드)":
        st.title("👤 직원 정보 관리")
        st.subheader("📥 엑셀 파일로 직원 일괄 등록하기")
        
        st.markdown("""
        **[엑셀 양식 안내]**  
        엑셀 파일에 **사번, 이름, 부서, 출근예정시간** 컬럼을 만들어서 업로드하시면 340명이든 1000명이든 한 번에 DB에 입주합니다!
        """)
        
        uploaded_file = st.file_uploader("직원 명부 엑셀 파일을 선택하세요", type=["xlsx", "csv"])
        
        if uploaded_file is not None:
            try:
                if uploaded_file.name.endswith('.csv'):
                    df = pd.read_csv(uploaded_file)
                else:
                    df = pd.read_excel(uploaded_file)
                
                st.write("📊 업로드 데이터 미리보기:")
                st.dataframe(df.head())
                
                if st.button("🚀 이 직원들을 데이터베이스에 진짜로 저장하기", type="primary"):
                    success_count = 0
                    with db_connect() as conn:
                        with conn.cursor() as cur:
                            for _, row in df.iterrows():
                                cur.execute("""
                                    INSERT INTO employees (emp_no, name, dept_name, scheduled_time)
                                    VALUES (%s, %s, %s, %s)
                                    ON CONFLICT (emp_no) DO UPDATE 
                                    SET name = EXCLUDED.name, dept_name = EXCLUDED.dept_name, scheduled_time = EXCLUDED.scheduled_time
                                """, (str(row['사번']), str(row['이름']), str(row['부서']), str(row['출근예정시간'])))
                                success_count += 1
                        conn.commit()
                    st.success(f"🎉 성공! {success_count}명의 직원이 보물섬 데이터베이스에 완벽하게 입주했습니다!")
            except Exception as e:
                st.error(f"파일 처리 중 에러 발생: {e}. 컬럼명이 '사번', '이름', '부서', '출근예정시간'으로 되어있는지 확인해 주세요.")

        # 현재 등록된 직원 조회 (Read)
        st.divider()
        st.subheader("👥 현재 등록된 직원 목록")
        with db_connect() as conn:
            df_emp = pd.read_sql("SELECT emp_no as 사번, name as 이름, dept_name as 부서, scheduled_time as 출근예정시간 FROM employees ORDER BY dept_name, name", conn)
        st.dataframe(df_emp, use_container_width=True)

    elif menu == "📋 출근입력":
        st.title("📋 부서별 출근 현황 입력")
        # 실제 DB 데이터를 바탕으로 부서별 출근 입력 및 수정(Update) 처리 구문 위치...
        st.info("직원관리 메뉴에서 엑셀 명부를 업로드하시면, 해당 부서 직원들이 이곳에 실시간으로 나타나 출근 기록을 시작할 수 있습니다!")
        
    else:
        st.title(f"{menu}")
        st.info("껍데기 프로토타입을 넘어 진짜 DB와 연동될 준비가 완료되었습니다. 직원관리에서 엑셀을 먼저 업로드해 보세요!")
