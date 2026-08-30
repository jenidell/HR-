"""
db.py
대교통신 사내 HR 플랫폼 - 데이터베이스 계층

지금은 SQLite로 동작합니다 (파일 하나로 끝나서 개발/소규모 운영에 충분).
70명 규모로 정식 운영하거나 Streamlit Community Cloud처럼 파일이
초기화될 수 있는 환경에 올릴 때는 Supabase(Postgres) 등으로 교체를
권장합니다. 이 파일의 함수 인터페이스만 유지하면 앱 코드(app.py)는
거의 손대지 않고 DB만 바꿀 수 있도록 설계했습니다.
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

DEPARTMENTS = ["관리팀", "소매팀", "전산팀", "도매팀", "재고팀", "기타"]


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
            department TEXT NOT NULL,
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

    # 최초 실행 시 관리자 계정이 하나도 없으면 기본 관리자 계정 생성
    cur.execute("SELECT COUNT(*) as cnt FROM users WHERE role = 'admin'")
    if cur.fetchone()["cnt"] == 0:
        create_user(
            username="admin",
            password="changeme123",
            name="관리자",
            department="관리팀",
            role="admin",
            _conn=conn,
        )
    conn.close()


def _hash_password(password: str, salt: str) -> str:
    return hashlib.pbkdf2_hmac(
        "sha256", password.encode("utf-8"), salt.encode("utf-8"), 100_000
    ).hex()


def create_user(username, password, name, department, role="employee", _conn=None):
    conn = _conn or get_conn()
    salt = os.urandom(16).hex()
    pw_hash = _hash_password(password, salt)
    now = datetime.now().isoformat()
    try:
        conn.execute(
            """INSERT INTO users (username, password_hash, salt, name, department, role, active, created_at)
               VALUES (?, ?, ?, ?, ?, ?, 1, ?)""",
            (username, pw_hash, salt, name, department, role, now),
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
            "SELECT id, username, name, department, role, active FROM users ORDER BY department, name"
        ).fetchall()
    else:
        rows = conn.execute(
            "SELECT id, username, name, department, role, active FROM users WHERE active = 1 ORDER BY department, name"
        ).fetchall()
    conn.close()
    return [dict(r) for r in rows]


def deactivate_user(user_id: int):
    conn = get_conn()
    conn.execute("UPDATE users SET active = 0 WHERE id = ?", (user_id,))
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


def get_all_attendance(start_date: str = None, end_date: str = None, department: str = None):
    conn = get_conn()
    q = """
        SELECT a.work_date, a.code, a.memo, u.name, u.department, u.username, a.updated_at
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
    if department and department != "전체":
        q += " AND u.department = ?"
        params.append(department)
    q += " ORDER BY a.work_date DESC, u.department, u.name"
    rows = conn.execute(q, params).fetchall()
    conn.close()
    return [dict(r) for r in rows]
