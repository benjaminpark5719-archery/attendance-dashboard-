"""
=================================================================
보물섬수산 출근현황 관리 - 데이터베이스 모듈 (v2.0 Final)
CSV 자동 마이그레이션 내장 / SQLAlchemy ORM
=================================================================
"""

import os
import csv
import hashlib
import secrets
from datetime import datetime, date
from pathlib import Path
from contextlib import contextmanager

from sqlalchemy import (
    create_engine, Column, Integer, String, Boolean, Date, DateTime,
    Float, ForeignKey, Text, UniqueConstraint, event
)
from sqlalchemy.orm import declarative_base, sessionmaker, relationship
import streamlit as st

Base = declarative_base()

# ─── Models ──────────────────────────────────────────────────────

class User(Base):
    __tablename__ = 'users'
    id = Column(Integer, primary_key=True, autoincrement=True)
    username = Column(String(50), unique=True, nullable=False)
    password_hash = Column(String(256), nullable=False)
    salt = Column(String(64), nullable=False)
    role = Column(String(20), nullable=False)
    display_name = Column(String(50), nullable=False)
    department_id = Column(Integer, ForeignKey('departments.id'), nullable=True)
    is_active = Column(Boolean, default=True)
    must_change_pw = Column(Boolean, default=True)
    created_at = Column(DateTime, default=datetime.now)
    updated_at = Column(DateTime, default=datetime.now, onupdate=datetime.now)
    department = relationship("Department", back_populates="users")


class Department(Base):
    __tablename__ = 'departments'
    id = Column(Integer, primary_key=True, autoincrement=True)
    name = Column(String(50), unique=True, nullable=False)
    code = Column(String(20), unique=True, nullable=False)
    sort_order = Column(Integer, default=0)
    is_active = Column(Boolean, default=True)
    created_at = Column(DateTime, default=datetime.now)
    updated_at = Column(DateTime, default=datetime.now, onupdate=datetime.now)
    users = relationship("User", back_populates="department")
    employees = relationship("Employee", back_populates="department")


class Employee(Base):
    __tablename__ = 'employees'
    id = Column(Integer, primary_key=True, autoincrement=True)
    emp_number = Column(String(20), unique=True, nullable=False)
    name = Column(String(50), nullable=False)
    department_id = Column(Integer, ForeignKey('departments.id'), nullable=False)
    position = Column(String(30), default='')
    phone = Column(String(20), default='')
    scheduled_time = Column(String(5), default='09:00')
    hire_date = Column(Date, nullable=True)
    resign_date = Column(Date, nullable=True)
    is_active = Column(Boolean, default=True)
    created_at = Column(DateTime, default=datetime.now)
    updated_at = Column(DateTime, default=datetime.now, onupdate=datetime.now)
    department = relationship("Department", back_populates="employees")
    attendance_logs = relationship("AttendanceLog", back_populates="employee")
    salary_info = relationship("SalaryInfo", back_populates="employee")


class AttendanceLog(Base):
    __tablename__ = 'attendance_logs'
    id = Column(Integer, primary_key=True, autoincrement=True)
    log_date = Column(Date, nullable=False)
    employee_id = Column(Integer, ForeignKey('employees.id'), nullable=False)
    department_id = Column(Integer, ForeignKey('departments.id'), nullable=False)
    scheduled_time = Column(String(5), default='')
    actual_time = Column(String(5), default='')
    status = Column(String(20), default='')
    note = Column(Text, default='')
    registered_by = Column(Integer, ForeignKey('users.id'), nullable=True)
    created_at = Column(DateTime, default=datetime.now)
    updated_at = Column(DateTime, default=datetime.now, onupdate=datetime.now)
    __table_args__ = (
        UniqueConstraint('log_date', 'employee_id', name='uq_daily_attendance'),
    )
    employee = relationship("Employee", back_populates="attendance_logs")
    department = relationship("Department")
    registrar = relationship("User")


class SalaryInfo(Base):
    __tablename__ = 'salary_info'
    id = Column(Integer, primary_key=True, autoincrement=True)
    employee_id = Column(Integer, ForeignKey('employees.id'), nullable=False)
    base_salary = Column(Integer, default=0)
    position_allowance = Column(Integer, default=0)
    meal_allowance = Column(Integer, default=0)
    transport_allowance = Column(Integer, default=0)
    overtime_allowance = Column(Integer, default=0)
    other_allowance = Column(Integer, default=0)
    national_pension = Column(Integer, default=0)
    health_insurance = Column(Integer, default=0)
    employment_insurance = Column(Integer, default=0)
    income_tax = Column(Integer, default=0)
    local_tax = Column(Integer, default=0)
    other_deduction = Column(Integer, default=0)
    pay_type = Column(String(20), default='월급')
    effective_date = Column(Date, nullable=False)
    notes = Column(Text, default='')
    created_at = Column(DateTime, default=datetime.now)
    updated_at = Column(DateTime, default=datetime.now, onupdate=datetime.now)
    employee = relationship("Employee", back_populates="salary_info")

    @property
    def total_allowance(self):
        return (self.position_allowance + self.meal_allowance +
                self.transport_allowance + self.overtime_allowance +
                self.other_allowance)

    @property
    def total_deduction(self):
        return (self.national_pension + self.health_insurance +
                self.employment_insurance + self.income_tax +
                self.local_tax + self.other_deduction)

    @property
    def net_salary(self):
        return self.base_salary + self.total_allowance - self.total_deduction


# ─── Engine & Session ────────────────────────────────────────────

def get_database_url():
    try:
        return st.secrets["database"]["url"]
    except Exception:
        db_url = os.environ.get("DATABASE_URL")
        if db_url:
            return db_url
    return "sqlite:///attendance.db"


@st.cache_resource
def get_engine():
    url = get_database_url()
    if url.startswith("sqlite"):
        engine = create_engine(url, echo=False,
                               connect_args={"check_same_thread": False})
        @event.listens_for(engine, "connect")
        def set_sqlite_pragma(dbapi_conn, connection_record):
            cursor = dbapi_conn.cursor()
            cursor.execute("PRAGMA foreign_keys=ON")
            cursor.close()
    else:
        engine = create_engine(url, echo=False, pool_pre_ping=True,
                               pool_size=5, max_overflow=10)
    return engine


@st.cache_resource
def get_session_factory():
    return sessionmaker(bind=get_engine())


@contextmanager
def get_db():
    Session = get_session_factory()
    session = Session()
    try:
        yield session
        session.commit()
    except Exception:
        session.rollback()
        raise
    finally:
        session.close()


# ─── Password Utilities ─────────────────────────────────────────

def hash_password(password: str, salt: str) -> str:
    return hashlib.pbkdf2_hmac(
        'sha256', password.encode(), salt.encode(), 100_000
    ).hex()

def verify_password(password: str, salt: str, stored_hash: str) -> bool:
    return hash_password(password, salt) == stored_hash


# ─── DB Init + CSV Auto-Migration ────────────────────────────────

DEPT_CONFIG = [
    ("회계부", "ACC", 1), ("회주방", "KIT", 2),
    ("식당 홀", "HLL", 3), ("식당 주방", "RKT", 4),
    ("배송팀", "DLV", 5), ("물류팀", "LOG", 6),
    ("해썹가공공장", "HAC", 7), ("마트 입점팀", "MRT", 8),
]

def init_db():
    """테이블 생성 → 초기 관리자 → CSV 자동 마이그레이션"""
    engine = get_engine()
    Base.metadata.create_all(engine)

    with get_db() as session:
        # 1) 초기 총괄관리자
        if session.query(User).count() == 0:
            salt = secrets.token_hex(32)
            pw_hash = hash_password("admin1234!", salt)
            session.add(User(
                username="admin", password_hash=pw_hash, salt=salt,
                role="super_admin", display_name="총괄관리자",
                is_active=True, must_change_pw=True
            ))

        # 2) 초기 부서
        if session.query(Department).count() == 0:
            for name, code, order in DEPT_CONFIG:
                session.add(Department(name=name, code=code, sort_order=order))
            session.flush()

        # 3) CSV 자동 마이그레이션 (직원 0명일 때만 실행)
        if session.query(Employee).count() == 0:
            _migrate_csv(session)


def _migrate_csv(session):
    """attendance_data/attendance.csv → employees + attendance_logs 일괄 삽입"""
    csv_path = Path("attendance_data") / "attendance.csv"
    if not csv_path.exists():
        return

    # 부서 매핑
    dept_map = {}
    for dept in session.query(Department).all():
        dept_map[dept.name] = dept

    # CSV 읽기
    rows = []
    try:
        with open(csv_path, encoding='utf-8-sig') as f:
            reader = csv.DictReader(f)
            for row in reader:
                rows.append(row)
    except Exception:
        return

    if not rows:
        return

    # 부서별 직원 카운터
    dept_counters = {}
    emp_objects = {}  # name → Employee

    for row in rows:
        dept_name = row.get('부서', '').strip()
        emp_name = row.get('이름', '').strip()
        scheduled = row.get('출근예정시간', '09:00').strip()

        if not dept_name or not emp_name:
            continue
        if dept_name not in dept_map:
            continue
        if emp_name in emp_objects:
            continue  # 이미 생성된 직원 스킵

        dept = dept_map[dept_name]
        dept_counters[dept.code] = dept_counters.get(dept.code, 0) + 1
        seq = dept_counters[dept.code]
        emp_number = f"{dept.code}-{seq:03d}"

        emp = Employee(
            emp_number=emp_number,
            name=emp_name,
            department_id=dept.id,
            position='',
            phone='',
            scheduled_time=scheduled or '09:00',
            hire_date=date.today(),
            is_active=True,
        )
        session.add(emp)
        emp_objects[emp_name] = emp

    session.flush()  # ID 확정

    # 출근 기록 마이그레이션
    for row in rows:
        emp_name = row.get('이름', '').strip()
        dept_name = row.get('부서', '').strip()
        if emp_name not in emp_objects or dept_name not in dept_map:
            continue

        emp = emp_objects[emp_name]
        dept = dept_map[dept_name]

        log_date_str = row.get('날짜', '').strip()
        try:
            log_date = datetime.strptime(log_date_str, '%Y-%m-%d').date()
        except (ValueError, TypeError):
            log_date = date.today()

        actual = row.get('실제출근시간', '').strip()
        status = row.get('근무상태', '').strip()
        note = row.get('비고', '').strip()

        if status:  # 상태가 있는 기록만
            session.add(AttendanceLog(
                log_date=log_date,
                employee_id=emp.id,
                department_id=dept.id,
                scheduled_time=emp.scheduled_time,
                actual_time=actual,
                status=status,
                note=note,
            ))


# ─── CRUD: Users ─────────────────────────────────────────────────

def authenticate_user(username: str, password: str):
    with get_db() as session:
        user = session.query(User).filter_by(
            username=username, is_active=True
        ).first()
        if user and verify_password(password, user.salt, user.password_hash):
            return {
                'id': user.id, 'username': user.username,
                'role': user.role, 'display_name': user.display_name,
                'department_id': user.department_id,
                'must_change_pw': user.must_change_pw
            }
    return None


def create_user(username, password, role, display_name, department_id=None):
    with get_db() as session:
        if session.query(User).filter_by(username=username).first():
            return False, "이미 존재하는 사용자명입니다."
        salt = secrets.token_hex(32)
        pw_hash = hash_password(password, salt)
        session.add(User(
            username=username, password_hash=pw_hash, salt=salt,
            role=role, display_name=display_name,
            department_id=department_id, must_change_pw=True
        ))
        return True, "사용자가 생성되었습니다."


def change_password(user_id, new_password):
    with get_db() as session:
        user = session.query(User).get(user_id)
        if user:
            salt = secrets.token_hex(32)
            user.password_hash = hash_password(new_password, salt)
            user.salt = salt
            user.must_change_pw = False
            return True
    return False


def get_all_users():
    with get_db() as session:
        users = session.query(User).order_by(User.id).all()
        return [{
            'id': u.id, 'username': u.username, 'role': u.role,
            'display_name': u.display_name,
            'department_id': u.department_id,
            'department_name': u.department.name if u.department else '-',
            'is_active': u.is_active,
        } for u in users]


def toggle_user_active(user_id, is_active):
    with get_db() as session:
        user = session.query(User).get(user_id)
        if user:
            user.is_active = is_active
            return True
    return False


def reset_user_password(user_id, new_password):
    with get_db() as session:
        user = session.query(User).get(user_id)
        if user:
            salt = secrets.token_hex(32)
            user.password_hash = hash_password(new_password, salt)
            user.salt = salt
            user.must_change_pw = True
            return True
    return False


# ─── CRUD: Departments ──────────────────────────────────────────

def get_all_departments(active_only=False):
    with get_db() as session:
        q = session.query(Department).order_by(Department.sort_order)
        if active_only:
            q = q.filter_by(is_active=True)
        depts = q.all()
        return [{
            'id': d.id, 'name': d.name, 'code': d.code,
            'sort_order': d.sort_order, 'is_active': d.is_active,
            'employee_count': session.query(Employee).filter_by(
                department_id=d.id, is_active=True).count()
        } for d in depts]


def create_department(name, code, sort_order=0):
    with get_db() as session:
        if session.query(Department).filter_by(name=name).first():
            return False, "이미 존재하는 부서명입니다."
        if session.query(Department).filter_by(code=code).first():
            return False, "이미 존재하는 부서코드입니다."
        session.add(Department(name=name, code=code, sort_order=sort_order))
        return True, "부서가 등록되었습니다."


def update_department(dept_id, name, code, sort_order):
    with get_db() as session:
        dept = session.query(Department).get(dept_id)
        if not dept:
            return False, "부서를 찾을 수 없습니다."
        dup = session.query(Department).filter(
            Department.name == name, Department.id != dept_id).first()
        if dup:
            return False, "이미 존재하는 부서명입니다."
        dup2 = session.query(Department).filter(
            Department.code == code, Department.id != dept_id).first()
        if dup2:
            return False, "이미 존재하는 부서코드입니다."
        dept.name = name
        dept.code = code
        dept.sort_order = sort_order
        return True, "부서 정보가 수정되었습니다."


def delete_department(dept_id):
    with get_db() as session:
        dept = session.query(Department).get(dept_id)
        if not dept:
            return False, "부서를 찾을 수 없습니다."
        emp_count = session.query(Employee).filter_by(
            department_id=dept_id, is_active=True).count()
        if emp_count > 0:
            return False, f"소속 직원 {emp_count}명이 있어 삭제할 수 없습니다."
        dept.is_active = False
        return True, "부서가 비활성화되었습니다."


# ─── CRUD: Employees ─────────────────────────────────────────────

def get_employees(department_id=None, active_only=True):
    with get_db() as session:
        q = session.query(Employee).join(Department)
        if department_id:
            q = q.filter(Employee.department_id == department_id)
        if active_only:
            q = q.filter(Employee.is_active == True)
        emps = q.order_by(Department.sort_order, Employee.emp_number).all()
        return [{
            'id': e.id, 'emp_number': e.emp_number, 'name': e.name,
            'department_id': e.department_id,
            'department_name': e.department.name,
            'position': e.position, 'phone': e.phone,
            'scheduled_time': e.scheduled_time,
            'hire_date': e.hire_date, 'resign_date': e.resign_date,
            'is_active': e.is_active,
        } for e in emps]


def create_employee(emp_number, name, department_id, position='',
                    phone='', scheduled_time='09:00', hire_date=None):
    with get_db() as session:
        if session.query(Employee).filter_by(emp_number=emp_number).first():
            return False, "이미 존재하는 사원번호입니다."
        session.add(Employee(
            emp_number=emp_number, name=name,
            department_id=department_id, position=position,
            phone=phone, scheduled_time=scheduled_time,
            hire_date=hire_date
        ))
        return True, "직원이 등록되었습니다."


def update_employee(emp_id, **kwargs):
    with get_db() as session:
        emp = session.query(Employee).get(emp_id)
        if not emp:
            return False, "직원을 찾을 수 없습니다."
        if 'emp_number' in kwargs:
            dup = session.query(Employee).filter(
                Employee.emp_number == kwargs['emp_number'],
                Employee.id != emp_id).first()
            if dup:
                return False, "이미 존재하는 사원번호입니다."
        for key, val in kwargs.items():
            if hasattr(emp, key):
                setattr(emp, key, val)
        return True, "직원 정보가 수정되었습니다."


def deactivate_employee(emp_id, resign_date=None):
    with get_db() as session:
        emp = session.query(Employee).get(emp_id)
        if not emp:
            return False, "직원을 찾을 수 없습니다."
        emp.is_active = False
        emp.resign_date = resign_date or date.today()
        return True, f"{emp.name}님이 퇴사 처리되었습니다."


def reactivate_employee(emp_id):
    with get_db() as session:
        emp = session.query(Employee).get(emp_id)
        if emp:
            emp.is_active = True
            emp.resign_date = None
            return True, f"{emp.name}님이 재활성화되었습니다."
    return False, "직원을 찾을 수 없습니다."


# ─── CRUD: Attendance ────────────────────────────────────────────

def get_attendance(log_date, department_id=None):
    with get_db() as session:
        q = session.query(Employee).join(Department).filter(Employee.is_active == True)
        if department_id:
            q = q.filter(Employee.department_id == department_id)
        employees = q.order_by(Department.sort_order, Employee.emp_number).all()

        result = []
        for emp in employees:
            log = session.query(AttendanceLog).filter_by(
                log_date=log_date, employee_id=emp.id).first()
            result.append({
                'employee_id': emp.id,
                'emp_number': emp.emp_number,
                'name': emp.name,
                'department_id': emp.department_id,
                'department_name': emp.department.name,
                'position': emp.position,
                'scheduled_time': emp.scheduled_time,
                'actual_time': log.actual_time if log else '',
                'status': log.status if log else '',
                'note': log.note if log else '',
            })
        return result


def save_attendance(log_date, records, registered_by_id=None):
    with get_db() as session:
        for rec in records:
            existing = session.query(AttendanceLog).filter_by(
                log_date=log_date, employee_id=rec['employee_id']).first()
            if existing:
                existing.actual_time = rec.get('actual_time', '')
                existing.status = rec.get('status', '')
                existing.note = rec.get('note', '')
                existing.registered_by = registered_by_id
                existing.updated_at = datetime.now()
            else:
                session.add(AttendanceLog(
                    log_date=log_date,
                    employee_id=rec['employee_id'],
                    department_id=rec['department_id'],
                    scheduled_time=rec.get('scheduled_time', ''),
                    actual_time=rec.get('actual_time', ''),
                    status=rec.get('status', ''),
                    note=rec.get('note', ''),
                    registered_by=registered_by_id
                ))
    return True


def get_attendance_stats(log_date):
    with get_db() as session:
        depts = session.query(Department).filter_by(
            is_active=True).order_by(Department.sort_order).all()
        stats = []
        total_all = {'총인원': 0, '출근완료': 0, '지각': 0, '결근': 0,
                     '휴무': 0, '조퇴': 0, '미입력': 0}
        for dept in depts:
            active_emps = session.query(Employee).filter_by(
                department_id=dept.id, is_active=True).count()
            if active_emps == 0:
                continue
            logs = session.query(AttendanceLog).filter_by(
                log_date=log_date, department_id=dept.id).all()
            sc = {}
            for log in logs:
                s = log.status if log.status else '미입력'
                sc[s] = sc.get(s, 0) + 1
            not_logged = active_emps - len(logs)
            row = {
                '부서': dept.name, '총인원': active_emps,
                '출근완료': sc.get('출근완료', 0), '지각': sc.get('지각', 0),
                '결근': sc.get('결근', 0), '휴무': sc.get('휴무', 0),
                '조퇴': sc.get('조퇴', 0),
                '미입력': not_logged + sc.get('', 0) + sc.get('미입력', 0),
            }
            working = row['총인원'] - row['휴무']
            row['출근율'] = f"{round(row['출근완료'] / working * 100, 1)}%" if working > 0 else "0%"
            stats.append(row)
            for k in total_all:
                total_all[k] += row[k]
        working_total = total_all['총인원'] - total_all['휴무']
        total_all['부서'] = '합계'
        total_all['출근율'] = f"{round(total_all['출근완료'] / working_total * 100, 1)}%" if working_total > 0 else "0%"
        stats.append(total_all)
        return stats


def get_attendance_history(employee_id=None, department_id=None,
                           date_from=None, date_to=None):
    with get_db() as session:
        q = session.query(AttendanceLog).join(Employee).join(Department)
        if employee_id:
            q = q.filter(AttendanceLog.employee_id == employee_id)
        if department_id:
            q = q.filter(AttendanceLog.department_id == department_id)
        if date_from:
            q = q.filter(AttendanceLog.log_date >= date_from)
        if date_to:
            q = q.filter(AttendanceLog.log_date <= date_to)
        logs = q.order_by(AttendanceLog.log_date.desc(), Department.sort_order).all()
        return [{
            'id': l.id, 'log_date': l.log_date,
            'emp_number': l.employee.emp_number,
            'name': l.employee.name,
            'department': l.department.name,
            'position': l.employee.position,
            'scheduled_time': l.scheduled_time,
            'actual_time': l.actual_time,
            'status': l.status, 'note': l.note,
            'registered_by': l.registrar.display_name if l.registrar else '-',
        } for l in logs]


# ─── CRUD: Salary ────────────────────────────────────────────────

def get_salary_info(employee_id):
    with get_db() as session:
        info = session.query(SalaryInfo).filter_by(
            employee_id=employee_id
        ).order_by(SalaryInfo.effective_date.desc()).first()
        if info:
            return {
                'id': info.id, 'employee_id': info.employee_id,
                'base_salary': info.base_salary,
                'position_allowance': info.position_allowance,
                'meal_allowance': info.meal_allowance,
                'transport_allowance': info.transport_allowance,
                'overtime_allowance': info.overtime_allowance,
                'other_allowance': info.other_allowance,
                'national_pension': info.national_pension,
                'health_insurance': info.health_insurance,
                'employment_insurance': info.employment_insurance,
                'income_tax': info.income_tax,
                'local_tax': info.local_tax,
                'other_deduction': info.other_deduction,
                'pay_type': info.pay_type,
                'effective_date': info.effective_date,
                'notes': info.notes,
                'total_allowance': info.total_allowance,
                'total_deduction': info.total_deduction,
                'net_salary': info.net_salary,
            }
    return None


def save_salary_info(employee_id, data):
    with get_db() as session:
        session.add(SalaryInfo(
            employee_id=employee_id,
            base_salary=data.get('base_salary', 0),
            position_allowance=data.get('position_allowance', 0),
            meal_allowance=data.get('meal_allowance', 0),
            transport_allowance=data.get('transport_allowance', 0),
            overtime_allowance=data.get('overtime_allowance', 0),
            other_allowance=data.get('other_allowance', 0),
            national_pension=data.get('national_pension', 0),
            health_insurance=data.get('health_insurance', 0),
            employment_insurance=data.get('employment_insurance', 0),
            income_tax=data.get('income_tax', 0),
            local_tax=data.get('local_tax', 0),
            other_deduction=data.get('other_deduction', 0),
            pay_type=data.get('pay_type', '월급'),
            effective_date=data.get('effective_date', date.today()),
            notes=data.get('notes', ''),
        ))
    return True


def get_salary_summary(department_id=None):
    with get_db() as session:
        q = session.query(Employee).join(Department).filter(Employee.is_active == True)
        if department_id:
            q = q.filter(Employee.department_id == department_id)
        employees = q.order_by(Department.sort_order, Employee.emp_number).all()
        result = []
        for emp in employees:
            latest = session.query(SalaryInfo).filter_by(
                employee_id=emp.id
            ).order_by(SalaryInfo.effective_date.desc()).first()
            result.append({
                'id': emp.id, 'emp_number': emp.emp_number,
                'name': emp.name, 'department': emp.department.name,
                'position': emp.position,
                'base_salary': latest.base_salary if latest else 0,
                'total_allowance': latest.total_allowance if latest else 0,
                'total_deduction': latest.total_deduction if latest else 0,
                'net_salary': latest.net_salary if latest else 0,
                'pay_type': latest.pay_type if latest else '-',
            })
        return result
