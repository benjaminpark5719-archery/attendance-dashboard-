-- =====================================================
-- 보물섬수산 HR v2.1 — Supabase SQL Editor에 붙여넣고 Run
-- =====================================================

-- 1. 직원 테이블 컬럼 추가 (25개 항목)
ALTER TABLE employees ADD COLUMN IF NOT EXISTS position VARCHAR(30) DEFAULT '';
ALTER TABLE employees ADD COLUMN IF NOT EXISTS birth_date VARCHAR(20) DEFAULT '';
ALTER TABLE employees ADD COLUMN IF NOT EXISTS gender VARCHAR(10) DEFAULT '';
ALTER TABLE employees ADD COLUMN IF NOT EXISTS phone VARCHAR(20) DEFAULT '';
ALTER TABLE employees ADD COLUMN IF NOT EXISTS emergency_contact VARCHAR(30) DEFAULT '';
ALTER TABLE employees ADD COLUMN IF NOT EXISTS address VARCHAR(200) DEFAULT '';
ALTER TABLE employees ADD COLUMN IF NOT EXISTS contract_type VARCHAR(20) DEFAULT '정규직';
ALTER TABLE employees ADD COLUMN IF NOT EXISTS contract_end_date VARCHAR(20) DEFAULT '';
ALTER TABLE employees ADD COLUMN IF NOT EXISTS hire_date VARCHAR(20) DEFAULT '';
ALTER TABLE employees ADD COLUMN IF NOT EXISTS resign_date VARCHAR(20) DEFAULT '';
ALTER TABLE employees ADD COLUMN IF NOT EXISTS employment_status VARCHAR(20) DEFAULT '재직';
ALTER TABLE employees ADD COLUMN IF NOT EXISTS base_salary INTEGER DEFAULT 0;
ALTER TABLE employees ADD COLUMN IF NOT EXISTS allowance INTEGER DEFAULT 0;
ALTER TABLE employees ADD COLUMN IF NOT EXISTS pay_type VARCHAR(20) DEFAULT '월급';
ALTER TABLE employees ADD COLUMN IF NOT EXISTS bank_name VARCHAR(30) DEFAULT '';
ALTER TABLE employees ADD COLUMN IF NOT EXISTS account_number VARCHAR(30) DEFAULT '';
ALTER TABLE employees ADD COLUMN IF NOT EXISTS evaluation_score VARCHAR(10) DEFAULT '';
ALTER TABLE employees ADD COLUMN IF NOT EXISTS awards TEXT DEFAULT '';
ALTER TABLE employees ADD COLUMN IF NOT EXISTS disciplines TEXT DEFAULT '';
ALTER TABLE employees ADD COLUMN IF NOT EXISTS health_cert_expiry VARCHAR(20) DEFAULT '';
ALTER TABLE employees ADD COLUMN IF NOT EXISTS safety_training_date VARCHAR(20) DEFAULT '';
ALTER TABLE employees ADD COLUMN IF NOT EXISTS memo TEXT DEFAULT '';

-- 2. TBM 기록 테이블
CREATE TABLE IF NOT EXISTS tbm_records (
    id SERIAL PRIMARY KEY,
    tbm_date DATE NOT NULL,
    department VARCHAR(50) NOT NULL,
    title VARCHAR(200) NOT NULL,
    content TEXT DEFAULT '',
    photo_data TEXT DEFAULT '',
    created_by VARCHAR(50) DEFAULT 'admin',
    created_at TIMESTAMP DEFAULT NOW()
);

-- 3. TBM 직원 확인(서명) 테이블
CREATE TABLE IF NOT EXISTS tbm_confirmations (
    id SERIAL PRIMARY KEY,
    tbm_id INTEGER NOT NULL,
    emp_no VARCHAR(20) NOT NULL,
    emp_name VARCHAR(50) DEFAULT '',
    confirmed_at TIMESTAMP DEFAULT NOW(),
    UNIQUE(tbm_id, emp_no)
);

-- 4. 기존 직원 데이터 초기화 후 80명 삽입
DELETE FROM attendance;
DELETE FROM employees;

INSERT INTO employees (emp_no, name, dept_name, scheduled_time, position, employment_status, hire_date) VALUES
('ACC-01','회계_직원01','회계부','09:00','사원','재직','2024-03-01'),
('ACC-02','회계_직원02','회계부','09:00','사원','재직','2024-03-01'),
('ACC-03','회계_직원03','회계부','09:00','대리','재직','2022-01-15'),
('ACC-04','회계_직원04','회계부','09:00','사원','재직','2024-06-01'),
('ACC-05','회계_직원05','회계부','09:00','과장','재직','2020-03-01'),
('ACC-06','회계_직원06','회계부','09:00','사원','재직','2025-01-02'),
('ACC-07','회계_직원07','회계부','09:00','사원','재직','2025-01-02'),
('ACC-08','회계_직원08','회계부','09:00','주임','재직','2023-07-01'),
('ACC-09','회계_직원09','회계부','09:00','사원','재직','2024-09-01'),
('ACC-10','회계_직원10','회계부','09:00','사원','재직','2025-03-01'),
('KIT-01','회주방_직원01','회주방','08:00','사원','재직','2023-05-01'),
('KIT-02','회주방_직원02','회주방','08:00','사원','재직','2023-05-01'),
('KIT-03','회주방_직원03','회주방','08:00','주임','재직','2021-11-01'),
('KIT-04','회주방_직원04','회주방','08:00','사원','재직','2024-02-01'),
('KIT-05','회주방_직원05','회주방','08:00','사원','재직','2024-02-01'),
('KIT-06','회주방_직원06','회주방','08:00','반장','재직','2020-06-01'),
('KIT-07','회주방_직원07','회주방','08:00','사원','재직','2025-01-02'),
('KIT-08','회주방_직원08','회주방','08:00','사원','재직','2025-04-01'),
('KIT-09','회주방_직원09','회주방','08:00','사원','재직','2024-08-01'),
('KIT-10','회주방_직원10','회주방','08:00','사원','재직','2024-08-01'),
('HLL-01','식당홀_직원01','식당 홀','09:00','사원','재직','2023-03-01'),
('HLL-02','식당홀_직원02','식당 홀','09:00','사원','재직','2023-03-01'),
('HLL-03','식당홀_직원03','식당 홀','09:00','주임','재직','2021-09-01'),
('HLL-04','식당홀_직원04','식당 홀','09:00','사원','재직','2024-01-02'),
('HLL-05','식당홀_직원05','식당 홀','09:00','사원','재직','2024-01-02'),
('HLL-06','식당홀_직원06','식당 홀','09:00','사원','재직','2024-07-01'),
('HLL-07','식당홀_직원07','식당 홀','09:00','사원','재직','2025-02-01'),
('HLL-08','식당홀_직원08','식당 홀','09:00','반장','재직','2020-04-01'),
('HLL-09','식당홀_직원09','식당 홀','09:00','사원','재직','2025-02-01'),
('HLL-10','식당홀_직원10','식당 홀','09:00','사원','재직','2024-11-01'),
('RKT-01','식당주방_직원01','식당 주방','08:00','사원','재직','2023-06-01'),
('RKT-02','식당주방_직원02','식당 주방','08:00','사원','재직','2023-06-01'),
('RKT-03','식당주방_직원03','식당 주방','08:00','주임','재직','2021-04-01'),
('RKT-04','식당주방_직원04','식당 주방','08:00','사원','재직','2024-03-01'),
('RKT-05','식당주방_직원05','식당 주방','08:00','과장','재직','2019-11-01'),
('RKT-06','식당주방_직원06','식당 주방','08:00','사원','재직','2024-10-01'),
('RKT-07','식당주방_직원07','식당 주방','08:00','사원','재직','2025-01-02'),
('RKT-08','식당주방_직원08','식당 주방','08:00','사원','재직','2025-01-02'),
('RKT-09','식당주방_직원09','식당 주방','08:00','반장','재직','2020-08-01'),
('RKT-10','식당주방_직원10','식당 주방','08:00','사원','재직','2024-05-01'),
('DLV-01','배송_직원01','배송팀','07:00','사원','재직','2023-04-01'),
('DLV-02','배송_직원02','배송팀','07:00','사원','재직','2023-04-01'),
('DLV-03','배송_직원03','배송팀','07:00','주임','재직','2021-06-01'),
('DLV-04','배송_직원04','배송팀','07:00','사원','재직','2024-01-02'),
('DLV-05','배송_직원05','배송팀','07:00','과장','재직','2019-07-01'),
('DLV-06','배송_직원06','배송팀','07:00','사원','재직','2024-09-01'),
('DLV-07','배송_직원07','배송팀','07:00','사원','재직','2025-03-01'),
('DLV-08','배송_직원08','배송팀','07:00','사원','재직','2025-03-01'),
('DLV-09','배송_직원09','배송팀','07:00','반장','재직','2020-02-01'),
('DLV-10','배송_직원10','배송팀','07:00','사원','재직','2024-06-01'),
('LOG-01','물류_직원01','물류팀','08:00','사원','재직','2023-02-01'),
('LOG-02','물류_직원02','물류팀','08:00','사원','재직','2023-02-01'),
('LOG-03','물류_직원03','물류팀','08:00','대리','재직','2021-01-15'),
('LOG-04','물류_직원04','물류팀','08:00','사원','재직','2024-04-01'),
('LOG-05','물류_직원05','물류팀','08:00','사원','재직','2024-04-01'),
('LOG-06','물류_직원06','물류팀','08:00','과장','재직','2019-09-01'),
('LOG-07','물류_직원07','물류팀','08:00','사원','재직','2025-01-02'),
('LOG-08','물류_직원08','물류팀','08:00','사원','재직','2025-05-01'),
('LOG-09','물류_직원09','물류팀','08:00','반장','재직','2020-10-01'),
('LOG-10','물류_직원10','물류팀','08:00','사원','재직','2024-07-01'),
('HAC-01','해썹_직원01','해썹가공공장','08:00','사원','재직','2023-08-01'),
('HAC-02','해썹_직원02','해썹가공공장','08:00','사원','재직','2023-08-01'),
('HAC-03','해썹_직원03','해썹가공공장','08:00','주임','재직','2022-03-01'),
('HAC-04','해썹_직원04','해썹가공공장','08:00','사원','재직','2024-05-01'),
('HAC-05','해썹_직원05','해썹가공공장','08:00','사원','재직','2024-05-01'),
('HAC-06','해썹_직원06','해썹가공공장','08:00','반장','재직','2020-01-02'),
('HAC-07','해썹_직원07','해썹가공공장','08:00','사원','재직','2025-02-01'),
('HAC-08','해썹_직원08','해썹가공공장','08:00','사원','재직','2025-02-01'),
('HAC-09','해썹_직원09','해썹가공공장','08:00','과장','재직','2019-05-01'),
('HAC-10','해썹_직원10','해썹가공공장','08:00','사원','재직','2024-12-01'),
('MRT-01','마트_직원01','마트 입점팀','09:00','사원','재직','2023-09-01'),
('MRT-02','마트_직원02','마트 입점팀','09:00','사원','재직','2023-09-01'),
('MRT-03','마트_직원03','마트 입점팀','09:00','대리','재직','2022-06-01'),
('MRT-04','마트_직원04','마트 입점팀','09:00','사원','재직','2024-02-01'),
('MRT-05','마트_직원05','마트 입점팀','09:00','사원','재직','2024-02-01'),
('MRT-06','마트_직원06','마트 입점팀','09:00','과장','재직','2020-05-01'),
('MRT-07','마트_직원07','마트 입점팀','09:00','사원','재직','2025-01-02'),
('MRT-08','마트_직원08','마트 입점팀','09:00','사원','재직','2025-04-01'),
('MRT-09','마트_직원09','마트 입점팀','09:00','반장','재직','2020-12-01'),
('MRT-10','마트_직원10','마트 입점팀','09:00','사원','재직','2024-10-01');

-- 완료!
