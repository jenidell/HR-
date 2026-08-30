"""
db.py
(주)대교통신 사내 HR 플랫폼 - 데이터베이스 계층

지금은 SQLite로 동작합니다 (파일 하나로 끝나서 개발/소규모 운영에 충분).
70명 규모로 정식 운영하거나 Streamlit Community Cloud처럼 파일이
초기화될 수 있는 환경에 올릴 때는 Supabase(Postgres) 등으로 교체를
권장합니다. 이 파일의 함수 인터페이스만 유지하면 앱 코드(app.py)는
거의 손대지 않고 DB만 바꿀 수 있도록 설계했습니다.

직원 구분 체계:
    - 본사: department 컬럼에 부서명(관리팀/소매팀/전산팀/도매팀/재고팀/기타) 저장
    - 직영: department 컬럼에 직영점 매장명 저장
    - 소사장: department 컬럼에 소사장 매장명 저장
  (실제 근태 파일 구조: 본사근태 / 직영근태 / 소사장근태·소사장출퇴근현황 을 반영)
"""

import sqlite3
import hashlib
import hmac
import os
from datetime import datetime, date

DB_PATH = os.path.join(os.path.dirname(__file__), "hr_platform.db")

# 근태 코드 정의 (기존 본사/소사장 근태 자동화에서 쓰던 코드 체계를 참고)
ATTENDANCE_CODES = [
    "정상출근",
    "연차",
    "반차",
    "지각",
    "조퇴",
    "특근",
    "당직",
    "교육",
    "경조",
    "예비군",
    "무급",
    "개인용무",
    "휴무",
    "기타",
]

# 직원 구분 (본사 / 직영 / 소사장)
CATEGORIES = ["본사", "직영", "소사장"]

# 본사 소속 부서
DEPARTMENTS = ["관리팀", "소매팀", "전산팀", "도매팀", "재고팀", "기타"]

# 직영점 매장 목록
JIKYEONG_STORES = ["범계역직영점", "상동점", "의왕점", "덕천직영점", "하안점"]

# 소사장 매장 목록
SOSAJANG_STORES = [
    "대교대리점 군포역점",
    "대교대리점 독산점",
    "대교대리점 만안점",
    "대교대리점 범계점",
    "대교대리점 본점",
    "대교대리점 산본점",
    "대교대리점 수원역직영점",
    "대교대리점 의왕역점",
    "대교대리점 스마트직영점",
    "대교대리점 인덕원점",
    "대교대리점 평촌역점",
    "대교대리점 평촌학원가점",
]


def get_org_units(category: str):
    """구분(본사/직영/소사장)에 따른 소속(부서 또는 매장) 목록 반환"""
    if category == "직영":
        return JIKYEONG_STORES
    if category == "소사장":
        return SOSAJANG_STORES
    return DEPARTMENTS


def get_conn():
    conn = sqlite3.connect(DB_PATH)
    conn.row_factory = sqlite3.Row
    conn.execute("PRAGMA foreign_keys = ON")
    return conn


def init_db():
    conn = get_conn()
    cur = conn.cursor()
    cur.execute(
        """
        CREATE TABLE IF NOT EXISTS users (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            username TEXT UNIQUE NOT NULL,
            password_hash TEXT NOT NULL,
            salt TEXT NOT NULL,
            name TEXT NOT NULL,
            category TEXT NOT NULL DEFAULT '본사',  -- '본사' / '직영' / '소사장'
            department TEXT NOT NULL,               -- 본사: 부서명 / 직영·소사장: 매장명
            role TEXT NOT NULL DEFAULT 'employee',  -- 'employee' or 'admin'
            active INTEGER NOT NULL DEFAULT 1,
            created_at TEXT NOT NULL
        )
        """
    )
    cur.execute(
        """
        CREATE TABLE IF NOT EXISTS attendance (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            user_id INTEGER NOT NULL,
            work_date TEXT NOT NULL,
            code TEXT NOT NULL,
            memo TEXT,
            created_at TEXT NOT NULL,
            updated_at TEXT NOT NULL,
            FOREIGN KEY (user_id) REFERENCES users(id),
            UNIQUE(user_id, work_date)
        )
        """
    )
    conn.commit()

    # 예전 버전(구조 변경 전) DB에는 category 컬럼이 없을 수 있어 마이그레이션
    try:
        cur.execute("ALTER TABLE users ADD COLUMN category TEXT NOT NULL DEFAULT '본사'")
        conn.commit()
    except sqlite3.OperationalError:
        pass  # 이미 컬럼이 있으면 무시

    # 최초 실행 시 관리자 계정이 하나도 없으면 기본 관리자 계정 생성
    cur.execute("SELECT COUNT(*) as cnt FROM users WHERE role = 'admin'")
    if cur.fetchone()["cnt"] == 0:
        create_user(
            username="admin",
            password="changeme123",
            name="관리자",
            category="본사",
            department="관리팀",
            role="admin",
            _conn=conn,
        )
    conn.close()


def _hash_password(password: str, salt: str) -> str:
    return hashlib.pbkdf2_hmac(
        "sha256", password.encode("utf-8"), salt.encode("utf-8"), 100_000
    ).hex()


def create_user(username, password, name, category, department, role="employee", _conn=None):
    conn = _conn or get_conn()
    salt = os.urandom(16).hex()
    pw_hash = _hash_password(password, salt)
    now = datetime.now().isoformat()
    try:
        conn.execute(
            """INSERT INTO users (username, password_hash, salt, name, category, department, role, active, created_at)
               VALUES (?, ?, ?, ?, ?, ?, ?, 1, ?)""",
            (username, pw_hash, salt, name, category, department, role, now),
        )
        conn.commit()
        return True, "계정이 생성되었습니다."
    except sqlite3.IntegrityError:
        return False, "이미 존재하는 아이디입니다."
    finally:
        if _conn is None:
            conn.close()


def verify_user(username: str, password: str):
    """성공 시 user row(dict-like)를 반환, 실패 시 None"""
    conn = get_conn()
    row = conn.execute(
        "SELECT * FROM users WHERE username = ? AND active = 1", (username,)
    ).fetchone()
    conn.close()
    if row is None:
        return None
    expected = row["password_hash"]
    actual = _hash_password(password, row["salt"])
    if hmac.compare_digest(expected, actual):
        return dict(row)
    return None


def list_users(include_inactive=False):
    conn = get_conn()
    if include_inactive:
        rows = conn.execute(
            "SELECT id, username, name, category, department, role, active FROM users ORDER BY category, department, name"
        ).fetchall()
    else:
        rows = conn.execute(
            "SELECT id, username, name, category, department, role, active FROM users WHERE active = 1 ORDER BY category, department, name"
        ).fetchall()
    conn.close()
    return [dict(r) for r in rows]


def list_users_by_org(category: str, org_unit: str):
    """특정 구분(본사/직영/소사장) + 부서·매장에 속한 재직자 목록"""
    conn = get_conn()
    rows = conn.execute(
        "SELECT id, username, name, category, department, role FROM users "
        "WHERE active = 1 AND category = ? AND department = ? ORDER BY name",
        (category, org_unit),
    ).fetchall()
    conn.close()
    return [dict(r) for r in rows]


def get_attendance_for_date(work_date: str, category: str = None, org_unit: str = None):
    """특정 날짜에 이미 입력된 근태를 user_id -> {code, memo}로 반환 (일괄입력 화면에서 기존값 미리 채우는 용도)"""
    conn = get_conn()
    q = """
        SELECT a.user_id, a.code, a.memo
        FROM attendance a JOIN users u ON a.user_id = u.id
        WHERE a.work_date = ?
    """
    params = [work_date]
    if category:
        q += " AND u.category = ?"
        params.append(category)
    if org_unit:
        q += " AND u.department = ?"
        params.append(org_unit)
    rows = conn.execute(q, params).fetchall()
    conn.close()
    return {r["user_id"]: {"code": r["code"], "memo": r["memo"]} for r in rows}


def deactivate_user(user_id: int):
    conn = get_conn()
    conn.execute("UPDATE users SET active = 0 WHERE id = ?", (user_id,))
    conn.commit()
    conn.close()


def update_user(user_id: int, name: str, category: str, department: str, role: str, new_password: str = None):
    """계정 정보 수정 (이름/구분/부서·매장/권한). new_password를 주면 비밀번호도 함께 변경"""
    conn = get_conn()
    if new_password:
        salt = os.urandom(16).hex()
        pw_hash = _hash_password(new_password, salt)
        conn.execute(
            "UPDATE users SET name = ?, category = ?, department = ?, role = ?, password_hash = ?, salt = ? WHERE id = ?",
            (name, category, department, role, pw_hash, salt, user_id),
        )
    else:
        conn.execute(
            "UPDATE users SET name = ?, category = ?, department = ?, role = ? WHERE id = ?",
            (name, category, department, role, user_id),
        )
    conn.commit()
    conn.close()


def delete_user(user_id: int):
    """계정을 완전히 삭제 (해당 계정의 근태 입력 내역도 함께 삭제됨). 잘못 만든 계정 정리용."""
    conn = get_conn()
    conn.execute("DELETE FROM attendance WHERE user_id = ?", (user_id,))
    conn.execute("DELETE FROM users WHERE id = ?", (user_id,))
    conn.commit()
    conn.close()


def delete_attendance(user_id: int, work_date: str):
    """특정 직원의 특정 날짜 근태 입력을 삭제 (잘못 입력한 경우 등)"""
    conn = get_conn()
    conn.execute(
        "DELETE FROM attendance WHERE user_id = ? AND work_date = ?",
        (user_id, work_date),
    )
    conn.commit()
    conn.close()


def upsert_attendance(user_id: int, work_date: str, code: str, memo: str = ""):
    """같은 사람/같은 날짜면 덮어쓰기 (수정), 아니면 새로 입력"""
    conn = get_conn()
    now = datetime.now().isoformat()
    existing = conn.execute(
        "SELECT id FROM attendance WHERE user_id = ? AND work_date = ?",
        (user_id, work_date),
    ).fetchone()
    if existing:
        conn.execute(
            "UPDATE attendance SET code = ?, memo = ?, updated_at = ? WHERE id = ?",
            (code, memo, now, existing["id"]),
        )
    else:
        conn.execute(
            """INSERT INTO attendance (user_id, work_date, code, memo, created_at, updated_at)
               VALUES (?, ?, ?, ?, ?, ?)""",
            (user_id, work_date, code, memo, now, now),
        )
    conn.commit()
    conn.close()


def get_user_attendance(user_id: int, start_date: str = None, end_date: str = None):
    conn = get_conn()
    q = "SELECT * FROM attendance WHERE user_id = ?"
    params = [user_id]
    if start_date:
        q += " AND work_date >= ?"
        params.append(start_date)
    if end_date:
        q += " AND work_date <= ?"
        params.append(end_date)
    q += " ORDER BY work_date DESC"
    rows = conn.execute(q, params).fetchall()
    conn.close()
    return [dict(r) for r in rows]


def get_all_attendance(start_date: str = None, end_date: str = None, category: str = None, org_unit: str = None):
    conn = get_conn()
    q = """
        SELECT a.work_date, a.code, a.memo, u.name, u.category, u.department, u.username, a.updated_at
        FROM attendance a JOIN users u ON a.user_id = u.id
        WHERE 1=1
    """
    params = []
    if start_date:
        q += " AND a.work_date >= ?"
        params.append(start_date)
    if end_date:
        q += " AND a.work_date <= ?"
        params.append(end_date)
    if category and category != "전체":
        q += " AND u.category = ?"
        params.append(category)
    if org_unit and org_unit != "전체":
        q += " AND u.department = ?"
        params.append(org_unit)
    q += " ORDER BY a.work_date DESC, u.category, u.department, u.name"
    rows = conn.execute(q, params).fetchall()
    conn.close()
    return [dict(r) for r in rows]
