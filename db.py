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
    "휴가",
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
# 순서는 본사 근태 양식의 실제 배치 순서와 같게 맞춰둡니다 (엑셀 생성 시 이 순서로 채움)
DEPARTMENTS = ["관리팀", "소매팀", "도매팀", "재고팀", "전산팀", "기타"]

# 직영점 매장 목록
# 순서는 직영점 근태 양식의 실제 배치 순서와 같게 맞춰둡니다
JIKYEONG_STORES = ["범계역직영점", "의왕점", "상동점", "하안점", "덕천직영점"]

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


# 최초 실행 시 자동으로 만들어지는 시스템 관리자 계정.
# 사람이 아니라 앱 관리용 계정이므로 근태 집계에서는 항상 제외합니다.
SYSTEM_ACCOUNT_USERNAME = "admin"
_NOT_SYSTEM = f"username != '{SYSTEM_ACCOUNT_USERNAME}'"


def list_attendance_users(category=None, org_unit=None):
    """근태 대상 직원 목록 (시스템 관리자 계정 제외)"""
    conn = get_conn()
    q = ("SELECT id, username, name, category, department, role FROM users "
         f"WHERE active = 1 AND {_NOT_SYSTEM}")
    params = []
    if category:
        q += " AND category = ?"
        params.append(category)
    if org_unit:
        q += " AND department = ?"
        params.append(org_unit)
    # 카톡에 이름이 나온 순서(sort_order)를 우선하고, 아직 모르는 사람은 이름순으로 뒤에
    q += " ORDER BY category, department, CASE WHEN sort_order > 0 THEN sort_order ELSE 9999 END, name"
    rows = conn.execute(q, params).fetchall()
    conn.close()
    return [dict(r) for r in rows]


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
    cur.execute(
        """
        CREATE TABLE IF NOT EXISTS announcements (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            title TEXT NOT NULL,
            content TEXT NOT NULL,
            author_id INTEGER NOT NULL,
            created_at TEXT NOT NULL,
            FOREIGN KEY (author_id) REFERENCES users(id)
        )
        """
    )
    cur.execute(
        """
        CREATE TABLE IF NOT EXISTS settings (
            key TEXT PRIMARY KEY,
            value TEXT NOT NULL
        )
        """
    )
    # 소사장은 '사람'이 아니라 '매장' 단위로 근태를 관리하기 때문에 별도 테이블을 씁니다.
    #   open_code  : 소사장 근태 양식(출첵) 하루 2칸 중 왼쪽 = 출근보고
    #   close_code : 오른쪽 = 퇴근보고
    #   perf_code  : 소사장 실적보고 양식(실적_월) 하루 1칸
    cur.execute(
        """
        CREATE TABLE IF NOT EXISTS store_attendance (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            category TEXT NOT NULL DEFAULT '소사장',
            store TEXT NOT NULL,
            work_date TEXT NOT NULL,
            open_code TEXT,
            close_code TEXT,
            perf_code TEXT,
            memo TEXT,
            created_at TEXT NOT NULL,
            updated_at TEXT NOT NULL,
            UNIQUE(category, store, work_date)
        )
        """
    )
    # 카톡 AI 검토 결과 보관 (검토 → 확인/수정 → DB 반영 흐름)
    cur.execute(
        """
        CREATE TABLE IF NOT EXISTS kakao_reviews (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            category TEXT NOT NULL,
            target_month TEXT NOT NULL,
            result_json TEXT NOT NULL,
            status TEXT NOT NULL DEFAULT 'pending',
            created_at TEXT NOT NULL
        )
        """
    )
    conn.commit()

    # 예전 버전(구조 변경 전) DB에는 category 컬럼이 없을 수 있어 마이그레이션
    for ddl in [
        "ALTER TABLE users ADD COLUMN category TEXT NOT NULL DEFAULT '본사'",
        # 카톡 보고에 이름이 나온 순서. 근태 양식을 이 순서대로 채웁니다(0이면 아직 모름 → 이름순).
        "ALTER TABLE users ADD COLUMN sort_order INTEGER NOT NULL DEFAULT 0",
    ]:
        try:
            cur.execute(ddl)
            conn.commit()
        except sqlite3.OperationalError:
            pass  # 이미 컬럼이 있으면 무시

    # 예전에 '기타'로 저장된 기록은 원래 표현이 memo에 남아 있습니다.
    # (예: 카톡에 "한송이 : 휴가"라고 올라왔는데 그때는 '휴가' 항목이 없어서 기타로 저장됨)
    # 그 사이에 정식 항목으로 추가된 표현이면 제대로 된 코드로 올려줍니다.
    try:
        marks = ",".join("?" * len(ATTENDANCE_CODES))
        cur.execute(
            f"UPDATE attendance SET code = memo, memo = '' "
            f"WHERE code = '기타' AND TRIM(memo) IN ({marks})",
            [c for c in ATTENDANCE_CODES],
        )
        if cur.rowcount:
            print(f"[db] '기타'로 저장돼 있던 {cur.rowcount}건을 정식 근태 항목으로 정리했습니다.")
        conn.commit()
    except sqlite3.OperationalError:
        pass

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
        f"WHERE active = 1 AND {_NOT_SYSTEM} AND category = ? AND department = ? ORDER BY name",
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


def activate_user(user_id: int):
    """비활성화된 계정을 다시 활성화 (되살리기)"""
    conn = get_conn()
    conn.execute("UPDATE users SET active = 1 WHERE id = ?", (user_id,))
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


def get_org_unit_summary_for_date(work_date: str):
    """해당 날짜 기준 구분·부서/매장별 재직 인원수와 근태 입력된 인원수를 반환.
    (관리자 화면의 '미입력 매장/부서 확인'에서 사용)"""
    conn = get_conn()
    unit_counts = conn.execute(
        f"SELECT category, department, COUNT(*) as cnt FROM users WHERE active = 1 AND {_NOT_SYSTEM} "
        "GROUP BY category, department"
    ).fetchall()
    entered = conn.execute(
        """
        SELECT u.category, u.department, COUNT(DISTINCT a.user_id) as cnt
        FROM attendance a JOIN users u ON a.user_id = u.id
        WHERE a.work_date = ?
        GROUP BY u.category, u.department
        """,
        (work_date,),
    ).fetchall()
    conn.close()

    entered_map = {(r["category"], r["department"]): r["cnt"] for r in entered}
    result = []
    for r in unit_counts:
        key = (r["category"], r["department"])
        result.append({
            "category": r["category"],
            "department": r["department"],
            "employee_count": r["cnt"],
            "entered_count": entered_map.get(key, 0),
        })
    return result


def get_daily_entry_counts(category: str, org_unit: str, start_date: str, end_date: str):
    """구분·부서/매장 기준, 기간 내 날짜별 근태 입력 인원수(중복 없이) 반환. 팀장이 '언제 안 올렸는지' 확인하는 용도."""
    conn = get_conn()
    rows = conn.execute(
        """
        SELECT a.work_date, COUNT(DISTINCT a.user_id) as cnt
        FROM attendance a JOIN users u ON a.user_id = u.id
        WHERE u.category = ? AND u.department = ? AND a.work_date >= ? AND a.work_date <= ?
        GROUP BY a.work_date
        """,
        (category, org_unit, start_date, end_date),
    ).fetchall()
    conn.close()
    return {r["work_date"]: r["cnt"] for r in rows}


def create_announcement(title: str, content: str, author_id: int):
    conn = get_conn()
    now = datetime.now().isoformat()
    conn.execute(
        "INSERT INTO announcements (title, content, author_id, created_at) VALUES (?, ?, ?, ?)",
        (title, content, author_id, now),
    )
    conn.commit()
    conn.close()


def list_announcements():
    conn = get_conn()
    rows = conn.execute(
        """
        SELECT a.id, a.title, a.content, a.created_at, u.name as author_name
        FROM announcements a JOIN users u ON a.author_id = u.id
        ORDER BY a.created_at DESC
        """
    ).fetchall()
    conn.close()
    return [dict(r) for r in rows]


def delete_announcement(announcement_id: int):
    conn = get_conn()
    conn.execute("DELETE FROM announcements WHERE id = ?", (announcement_id,))
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


def get_attendance_records_for_matrix(category: str, start_date: str, end_date: str):
    """구분(본사/직영/소사장)별 근태표(매트릭스) 엑셀 생성용 - user_id 포함 근태 기록 반환.
    (기존 get_all_attendance는 user_id가 빠져있어 매트릭스 조립에 부족함)"""
    conn = get_conn()
    rows = conn.execute(
        """
        SELECT a.user_id, a.work_date, a.code
        FROM attendance a JOIN users u ON a.user_id = u.id
        WHERE u.category = ? AND a.work_date >= ? AND a.work_date <= ?
        """,
        (category, start_date, end_date),
    ).fetchall()
    conn.close()
    return [dict(r) for r in rows]


def get_setting(key: str, default: str = None):
    """앱 설정값(예: 탭 순서) 조회. 없으면 default 반환."""
    conn = get_conn()
    row = conn.execute("SELECT value FROM settings WHERE key = ?", (key,)).fetchone()
    conn.close()
    return row["value"] if row else default


def set_setting(key: str, value: str):
    """앱 설정값 저장(있으면 덮어쓰기)."""
    conn = get_conn()
    conn.execute(
        "INSERT INTO settings (key, value) VALUES (?, ?) "
        "ON CONFLICT(key) DO UPDATE SET value = excluded.value",
        (key, value),
    )
    conn.commit()
    conn.close()


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


# ============================================================================
# 소사장 매장 단위 근태 (출첵 / 실적_월)
# ============================================================================

# 소사장 근태 양식(출첵)에서 쓰는 코드
STORE_CHULCHECK_CODES = ["o", "휴", "미", "개인", "휴가", "본사", "조기마감", "조기퇴근"]
# 소사장 실적보고 양식(실적_월)에서 쓰는 코드
STORE_PERF_CODES = ["O", "휴", "미", "휴가", "개인", "본사", "조기마감", "조기퇴근"]


def upsert_store_attendance(store, work_date, open_code=None, close_code=None,
                            perf_code=None, memo=None, category="소사장"):
    """매장 단위 근태 저장 (같은 매장·같은 날짜면 덮어쓰기).
    None으로 넘긴 항목은 기존 값을 그대로 둡니다(부분 수정 가능)."""
    conn = get_conn()
    now = datetime.now().isoformat()
    row = conn.execute(
        "SELECT * FROM store_attendance WHERE category=? AND store=? AND work_date=?",
        (category, store, work_date),
    ).fetchone()
    if row:
        conn.execute(
            """UPDATE store_attendance
               SET open_code=?, close_code=?, perf_code=?, memo=?, updated_at=?
               WHERE id=?""",
            (
                open_code if open_code is not None else row["open_code"],
                close_code if close_code is not None else row["close_code"],
                perf_code if perf_code is not None else row["perf_code"],
                memo if memo is not None else row["memo"],
                now,
                row["id"],
            ),
        )
    else:
        conn.execute(
            """INSERT INTO store_attendance
               (category, store, work_date, open_code, close_code, perf_code, memo, created_at, updated_at)
               VALUES (?,?,?,?,?,?,?,?,?)""",
            (category, store, work_date, open_code, close_code, perf_code, memo, now, now),
        )
    conn.commit()
    conn.close()


def get_store_attendance(start_date, end_date, category="소사장"):
    """기간 내 매장 단위 근태 전체 조회"""
    conn = get_conn()
    rows = conn.execute(
        """SELECT store, work_date, open_code, close_code, perf_code, memo
           FROM store_attendance
           WHERE category=? AND work_date>=? AND work_date<=?
           ORDER BY store, work_date""",
        (category, start_date, end_date),
    ).fetchall()
    conn.close()
    return [dict(r) for r in rows]


def delete_store_attendance(store, work_date, category="소사장"):
    conn = get_conn()
    conn.execute(
        "DELETE FROM store_attendance WHERE category=? AND store=? AND work_date=?",
        (category, store, work_date),
    )
    conn.commit()
    conn.close()


def get_store_daily_entry_counts(start_date, end_date, category="소사장"):
    """날짜별로 근태가 들어온 매장 수 (반영 현황 캘린더용)"""
    conn = get_conn()
    rows = conn.execute(
        """SELECT work_date, COUNT(DISTINCT store) AS cnt
           FROM store_attendance
           WHERE category=? AND work_date>=? AND work_date<=?
           GROUP BY work_date""",
        (category, start_date, end_date),
    ).fetchall()
    conn.close()
    return {r["work_date"]: r["cnt"] for r in rows}


# ============================================================================
# 카톡 AI 검토 결과
# ============================================================================

def save_kakao_review(category, target_month, result_json, status="pending"):
    conn = get_conn()
    cur = conn.cursor()
    cur.execute(
        """INSERT INTO kakao_reviews (category, target_month, result_json, status, created_at)
           VALUES (?,?,?,?,?)""",
        (category, target_month, result_json, status, datetime.now().isoformat()),
    )
    review_id = cur.lastrowid
    conn.commit()
    conn.close()
    return review_id


def list_kakao_reviews(category=None, limit=20):
    conn = get_conn()
    q = "SELECT id, category, target_month, status, created_at FROM kakao_reviews"
    params = []
    if category:
        q += " WHERE category = ?"
        params.append(category)
    q += " ORDER BY created_at DESC LIMIT ?"
    params.append(limit)
    rows = conn.execute(q, params).fetchall()
    conn.close()
    return [dict(r) for r in rows]


def get_kakao_review(review_id):
    conn = get_conn()
    row = conn.execute("SELECT * FROM kakao_reviews WHERE id = ?", (review_id,)).fetchone()
    conn.close()
    return dict(row) if row else None


def update_kakao_review(review_id, result_json=None, status=None):
    conn = get_conn()
    if result_json is not None:
        conn.execute("UPDATE kakao_reviews SET result_json=? WHERE id=?", (result_json, review_id))
    if status is not None:
        conn.execute("UPDATE kakao_reviews SET status=? WHERE id=?", (status, review_id))
    conn.commit()
    conn.close()


def delete_kakao_review(review_id):
    conn = get_conn()
    conn.execute("DELETE FROM kakao_reviews WHERE id = ?", (review_id,))
    conn.commit()
    conn.close()


def find_user_by_name(name, category=None):
    """이름으로 직원 찾기 (AI 검토 결과를 DB에 반영할 때 사용). 동명이인이면 여러 건 반환."""
    conn = get_conn()
    q = "SELECT id, name, category, department FROM users WHERE active=1 AND name=?"
    params = [name]
    if category:
        q += " AND category=?"
        params.append(category)
    rows = conn.execute(q, params).fetchall()
    conn.close()
    return [dict(r) for r in rows]


def set_kakao_order(user_ids):
    """카톡 보고에 이름이 나온 순서를 기억합니다 (근태 양식을 이 순서로 채우기 위함).
    user_ids : 카톡에서 나온 순서대로의 직원 id 목록 (중복은 첫 등장만 반영)"""
    conn = get_conn()
    seen, n = set(), 0
    for uid in user_ids:
        if uid in seen:
            continue
        seen.add(uid)
        n += 1
        conn.execute("UPDATE users SET sort_order = ? WHERE id = ?", (n, uid))
    conn.commit()
    conn.close()
    return n
