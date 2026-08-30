"""
(주)대교통신 사내 HR 플랫폼 - Phase 1 (근태입력 셀프서비스 MVP)

실행:
    streamlit run app.py

기본 관리자 계정: admin / changeme123  (반드시 최초 로그인 후 비밀번호를 바꿀 직원용 계정을 새로 만들고,
                                       admin 계정 비밀번호도 바꿔서 사용하세요)
"""

import io
import os
import re
from datetime import date, timedelta

import pandas as pd
import streamlit as st
from openpyxl import Workbook

import db

st.set_page_config(page_title="(주)대교통신 HR 플랫폼", page_icon="🗂️", layout="wide")
db.init_db()

DOCS_DIR = os.path.join(os.path.dirname(__file__), "docs")


def _list_docs(subfolder: str):
    """docs/<subfolder> 안의 파일 목록을 (파일명, 전체경로) 리스트로 반환. 폴더가 없으면 빈 리스트."""
    folder = os.path.join(DOCS_DIR, subfolder)
    if not os.path.isdir(folder):
        return []
    files = []
    for fname in sorted(os.listdir(folder)):
        fpath = os.path.join(folder, fname)
        if os.path.isfile(fpath) and not fname.startswith(".") and fname.lower() != "readme.md":
            files.append((fname, fpath))
    return files


def documents_view(subfolder: str, empty_msg: str):
    files = _list_docs(subfolder)
    if not files:
        st.info(empty_msg)
        return
    for fname, fpath in files:
        with open(fpath, "rb") as f:
            data = f.read()
        st.download_button(
            f"📄 {fname}", data=data, file_name=fname,
            key=f"doc_dl_{subfolder}_{fname}", use_container_width=True,
        )


def _build_bulk_template():
    """직원 일괄 등록용 빈 엑셀 양식 (예시 3줄 포함) 생성"""
    sample = pd.DataFrame([
        {"아이디": "D100001", "이름": "홍길동", "구분": "본사", "부서/매장": "관리팀",
         "초기비밀번호": "changeme01", "권한": "employee"},
        {"아이디": "D100002", "이름": "김철수", "구분": "직영", "부서/매장": "범계역직영점",
         "초기비밀번호": "changeme02", "권한": "employee"},
        {"아이디": "D100003", "이름": "이영희", "구분": "소사장", "부서/매장": "대교대리점 군포역점",
         "초기비밀번호": "changeme03", "권한": "employee"},
    ])
    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as writer:
        sample.to_excel(writer, index=False, sheet_name="직원목록")
    return buf.getvalue()


# ---------------------------------------------------------------------------
# 로그인 / 세션 관리
# ---------------------------------------------------------------------------
def login_view():
    st.title("🗂️ (주)대교통신 HR 플랫폼")
    st.caption("근태입력 셀프서비스 (Phase 1)")

    with st.form("login_form"):
        username = st.text_input("아이디")
        password = st.text_input("비밀번호", type="password")
        submitted = st.form_submit_button("로그인", use_container_width=True)

    if submitted:
        user = db.verify_user(username.strip(), password)
        if user:
            st.session_state["user"] = user
            st.rerun()
        else:
            st.error("아이디 또는 비밀번호가 올바르지 않습니다.")


def logout_button():
    with st.sidebar:
        user = st.session_state["user"]
        st.markdown(f"**{user['name']}**")
        st.caption(f"{user['category']} · {user['department']}")
        st.caption("관리자" if user["role"] == "admin" else "직원")
        if st.button("로그아웃", use_container_width=True):
            del st.session_state["user"]
            st.rerun()


# ---------------------------------------------------------------------------
# 직원용: 근태입력
# ---------------------------------------------------------------------------
def employee_view(user):
    st.header("근태 입력")

    col1, col2 = st.columns([1, 1])
    with col1:
        with st.form("attendance_form"):
            work_date = st.date_input("날짜", value=date.today())
            code = st.selectbox("근태 코드", db.ATTENDANCE_CODES)
            memo = st.text_input("메모 (선택)", placeholder="예: 오전 반차, 결혼식 참석 등")
            submitted = st.form_submit_button("입력/수정 저장", use_container_width=True)
        if submitted:
            db.upsert_attendance(user["id"], work_date.isoformat(), code, memo)
            st.success(f"{work_date.isoformat()} 근태가 '{code}'(으)로 저장되었습니다.")
            st.rerun()

    with col2:
        st.markdown("**근태 코드 안내**")
        st.caption(
            "정상출근 · 연차 · 반차 · 지각 · 조퇴 · 특근 · 당직 · 교육 · "
            "경조 · 예비군 · 무급 · 개인용무 · 휴무 · 기타\n\n"
            "같은 날짜에 다시 입력하면 기존 내용을 덮어씁니다(수정)."
        )
        st.info(
            "💡 담당자가 '일괄입력'에서 같은 날짜를 저장하면 여기서 입력한 내용도 "
            "그 내용으로 덮어써요. 평소엔 담당자의 일괄입력을 기본으로 하고, "
            "개인 입력은 급하게 미리 남겨둘 때만 쓰는 걸 추천해요."
        )

    st.divider()
    st.subheader("내 최근 입력 내역")

    default_start = date.today() - timedelta(days=30)
    c1, c2 = st.columns(2)
    with c1:
        start = st.date_input("조회 시작일", value=default_start, key="emp_start")
    with c2:
        end = st.date_input("조회 종료일", value=date.today(), key="emp_end")

    records = db.get_user_attendance(user["id"], start.isoformat(), end.isoformat())
    if not records:
        st.info("해당 기간에 입력된 근태 내역이 없습니다.")
        return

    for r in records:
        with st.container(border=True):
            top_l, top_r = st.columns([2, 2])
            top_l.markdown(f"**{r['work_date']}**")
            top_r.caption(f"최종수정 {(r['updated_at'] or '')[:16].replace('T', ' ')}")

            code_idx = db.ATTENDANCE_CODES.index(r["code"]) if r["code"] in db.ATTENDANCE_CODES else 0
            c_code, c_memo = st.columns([1, 2])
            new_code = c_code.selectbox(
                "근태코드", db.ATTENDANCE_CODES, index=code_idx,
                key=f"hist_code_{r['work_date']}", label_visibility="collapsed",
            )
            new_memo = c_memo.text_input(
                "메모", value=r["memo"] or "", key=f"hist_memo_{r['work_date']}",
                label_visibility="collapsed", placeholder="메모",
            )

            b_edit, b_del = st.columns(2)
            if b_edit.button("수정 저장", key=f"hist_edit_{r['work_date']}", use_container_width=True):
                db.upsert_attendance(user["id"], r["work_date"], new_code, new_memo)
                st.success(f"{r['work_date']} 근태가 수정되었습니다.")
                st.rerun()
            if b_del.button("삭제", key=f"hist_del_{r['work_date']}", type="secondary", use_container_width=True):
                db.delete_attendance(user["id"], r["work_date"])
                st.success(f"{r['work_date']} 근태 입력이 삭제되었습니다.")
                st.rerun()


# ---------------------------------------------------------------------------
# 일괄입력: 본사 부서 / 직영·소사장 매장 담당자가 하루치 소속 인원 전체를 한번에 입력
# ---------------------------------------------------------------------------
def bulk_entry_view(user):
    st.header("일괄입력")
    st.caption("본사 부서, 직영·소사장 매장 담당자가 하루 단위로 소속 인원 전체 근태를 한번에 입력할 수 있어요.")
    st.caption(
        "💡 헷갈리지 않으려면: 이 화면을 '공식 입력'으로 쓰고, 아직 선택 안 한 사람은 "
        "그대로 두면 저장되지 않아요 — 정상출근이 자동으로 저장되지 않으니 안심하고 "
        "빈 채로 넘어가도 됩니다."
    )

    c1, c2, c3 = st.columns(3)
    with c1:
        work_date = st.date_input("날짜", value=date.today(), key="bulk_date")
    with c2:
        cat_options = db.CATEGORIES
        default_cat_idx = cat_options.index(user["category"]) if user["category"] in cat_options else 0
        category = st.selectbox("구분", cat_options, index=default_cat_idx, key="bulk_category")
    with c3:
        org_options = db.get_org_units(category)
        if category == user["category"] and user["department"] in org_options:
            default_org_idx = org_options.index(user["department"])
        else:
            default_org_idx = 0
        org_unit = st.selectbox(
            "부서/매장 선택", org_options, index=default_org_idx, key=f"bulk_org_{category}"
        )

    members = db.list_users_by_org(category, org_unit)
    if not members:
        st.info(f"'{org_unit}'에 소속된 직원이 없습니다.")
        return

    existing = db.get_attendance_for_date(work_date.isoformat(), category, org_unit)

    entered_count = len(existing)
    member_count = len(members)
    if work_date == date.today():
        if entered_count == 0:
            st.warning(f"⚠️ 오늘({work_date.isoformat()}) {org_unit} 근태를 아직 아무도 입력하지 않았어요.")
        elif entered_count < member_count:
            st.info(f"📝 오늘 {org_unit} {entered_count}/{member_count}명 입력 완료. 나머지 인원도 확인해주세요.")
        else:
            st.success(f"✅ 오늘 {org_unit} 전원({member_count}명) 입력 완료!")
    else:
        st.caption(f"선택한 날짜({work_date.isoformat()}) 기준 {org_unit} {entered_count}/{member_count}명 입력됨")

    with st.expander(f"{org_unit} 이번 달 날짜별 입력 현황 (언제 안 올렸는지 확인)"):
        month_start = work_date.replace(day=1)
        today = date.today()
        month_end = min(
            date(work_date.year, work_date.month + 1, 1) - timedelta(days=1)
            if work_date.month < 12 else date(work_date.year, 12, 31),
            today,
        )
        counts = db.get_daily_entry_counts(
            category, org_unit, month_start.isoformat(), month_end.isoformat()
        )
        weekday_kr = ["월", "화", "수", "목", "금", "토", "일"]
        rows = []
        d = month_start
        while d <= month_end:
            cnt = counts.get(d.isoformat(), 0)
            if cnt == 0:
                status = "❌ 미입력"
            elif cnt < member_count:
                status = f"⚠️ 일부만 ({cnt}/{member_count}명)"
            else:
                status = f"✅ 완료 ({cnt}/{member_count}명)"
            rows.append({"날짜": d.isoformat(), "요일": weekday_kr[d.weekday()], "상태": status})
            d += timedelta(days=1)
        st.dataframe(pd.DataFrame(rows), use_container_width=True, hide_index=True)

    NOT_SELECTED = "─ 선택 안 함 ─"
    BULK_CODE_OPTIONS = [NOT_SELECTED] + db.ATTENDANCE_CODES

    with st.form("bulk_attendance_form"):
        entries = {}
        for m in members:
            has_prev = m["id"] in existing
            prev = existing.get(m["id"], {})
            prev_code = prev.get("code", "")
            prev_memo = prev.get("memo", "")
            code_idx = BULK_CODE_OPTIONS.index(prev_code) if has_prev and prev_code in db.ATTENDANCE_CODES else 0

            col_name, col_code, col_memo = st.columns([2, 2, 3])
            with col_name:
                st.markdown(f"**{m['name']}**")
                if has_prev:
                    st.caption(f"✅ 입력됨 ({prev_code})")
                else:
                    st.caption("⏳ 미입력")
            with col_code:
                code = st.selectbox(
                    "근태 코드", BULK_CODE_OPTIONS, index=code_idx,
                    key=f"bulk_code_{m['id']}", label_visibility="collapsed",
                )
            with col_memo:
                memo = st.text_input(
                    "메모", value=prev_memo, key=f"bulk_memo_{m['id']}",
                    label_visibility="collapsed", placeholder="메모 (선택)",
                )
            entries[m["id"]] = (code, memo, m["name"])

        submitted = st.form_submit_button(f"{org_unit} 전체 저장 ({work_date.isoformat()})", use_container_width=True)

    if submitted:
        saved = 0
        skipped_names = []
        for user_id, (code, memo, name) in entries.items():
            if code == NOT_SELECTED:
                skipped_names.append(name)
                continue
            db.upsert_attendance(user_id, work_date.isoformat(), code, memo)
            saved += 1
        if saved:
            st.success(f"{work_date.isoformat()} 기준 {org_unit} {saved}명 근태가 저장되었습니다.")
        if skipped_names:
            st.warning(
                f"⏳ 근태를 선택하지 않아 저장하지 않은 인원 {len(skipped_names)}명: "
                f"{', '.join(skipped_names)} — 확인 후 다시 입력해주세요."
            )
        st.rerun()


# ---------------------------------------------------------------------------
# 공지사항
# ---------------------------------------------------------------------------
def announcements_view(user):
    st.header("공지사항")

    if user["role"] == "admin":
        with st.expander("공지 작성"):
            with st.form("new_announcement_form"):
                title = st.text_input("제목")
                content = st.text_area("내용", height=150)
                submitted = st.form_submit_button("게시", use_container_width=True)
            if submitted:
                if not title.strip() or not content.strip():
                    st.error("제목과 내용을 모두 입력해주세요.")
                else:
                    db.create_announcement(title.strip(), content.strip(), user["id"])
                    st.success("공지사항이 등록되었습니다.")
                    st.rerun()
        st.divider()

    announcements = db.list_announcements()
    if not announcements:
        st.info("등록된 공지사항이 없습니다.")
        return

    for a in announcements:
        with st.container(border=True):
            st.markdown(f"**{a['title']}**")
            st.caption(f"{a['author_name']} · {(a['created_at'] or '')[:16].replace('T', ' ')}")
            st.write(a["content"])
            if user["role"] == "admin":
                if st.button("삭제", key=f"ann_del_{a['id']}"):
                    db.delete_announcement(a["id"])
                    st.success("삭제되었습니다.")
                    st.rerun()


# ---------------------------------------------------------------------------
# 카톡 근태 가져오기: 본사/직영/소사장 카톡방에 매일 올라오는 근태 메시지를
# 그대로 붙여넣으면 근태표 형식으로 파싱 (직원들의 카톡 습관은 그대로 유지)
# ---------------------------------------------------------------------------

_KAKAO_STATUS_MAP = {
    "출근": "정상출근",
    "연차": "연차",
    "반차": "반차",
    "오전반차": "반차",
    "오후반차": "반차",
    "지각": "지각",
    "조퇴": "조퇴",
    "특근": "특근",
    "당직": "당직",
    "교육": "교육",
    "경조": "경조",
    "예비군": "예비군",
    "무급": "무급",
    "개인용무": "개인용무",
    "휴무": "휴무",
}

_DATE_DIVIDER_RE = re.compile(r"(\d{4})년\s*(\d{1,2})월\s*(\d{1,2})일")

# 카카오톡 "대화 내보내기" .txt 파일은 메시지 첫 줄이 "[보낸사람] [오전/오후 시:분] 내용" 형태로 시작함.
# 화면에서 복사해서 붙여넣을 땐 이 접두어가 없을 수도 있어서, 있으면 제거하고 없으면 그대로 둠.
_SENDER_PREFIX_RE = re.compile(r"^\[.+?\]\s*\[(?:오전|오후)\s*\d{1,2}:\d{2}\]\s*")

_STATUS_TIME_MEMO = {"오전반차": "오전", "오후반차": "오후"}


def _strip_kakao_prefix(raw_line):
    return _SENDER_PREFIX_RE.sub("", raw_line, count=1).strip()


def parse_headquarters_chat(text, fallback_date):
    """본사 카톡 형식: "8월 28일 (금) 전산팀 출근현황" 헤더 다음 "이름 : 상태" 줄들.
    "-이름 : 상태"처럼 앞에 - 가 붙는 경우, 카카오톡 대화 내보내기 [보낸사람] [시간] 접두어도 처리."""
    header_re = re.compile(r"(\d{1,2})월\s*(\d{1,2})일.*?(?:출근현황|근태)")
    name_status_re = re.compile(r"^\s*-?\s*([가-힣A-Za-z0-9]+)\s*[:：]\s*(.+?)\s*$")

    results = []
    current_date = fallback_date
    for raw in text.splitlines():
        line = _strip_kakao_prefix(raw)
        if not line:
            continue
        dm = _DATE_DIVIDER_RE.search(line)
        if dm:
            try:
                current_date = date(int(dm.group(1)), int(dm.group(2)), int(dm.group(3)))
            except ValueError:
                pass
            continue
        hm = header_re.search(line)
        if hm:
            month, day = int(hm.group(1)), int(hm.group(2))
            try:
                current_date = date(current_date.year, month, day)
            except ValueError:
                pass
            continue
        nm = name_status_re.match(line)
        if nm:
            name, status_text = nm.group(1), nm.group(2).strip()
            code = _KAKAO_STATUS_MAP.get(status_text)
            results.append({
                "raw": line, "name": name, "store": None,
                "work_date": current_date, "code": code or "기타",
                "memo": _STATUS_TIME_MEMO.get(status_text, "" if code else status_text),
            })
    return results


def parse_jikyeong_chat(text, fallback_date):
    """직영 카톡 형식: "매장명 출근보고" / "10시 이름1 이름2" / "휴무 이름A 이름B" / "이상입니다" """
    report_start_re = re.compile(r"^(.+?)\s*출근보고\s*$")
    time_line_re = re.compile(r"^\d{1,2}시\s*(.+)$")
    end_re = re.compile(r"^이상입니다\s*$")

    lines = [_strip_kakao_prefix(l) for l in text.splitlines()]
    results = []
    current_date = fallback_date
    i = 0
    while i < len(lines):
        line = lines[i]
        if not line:
            i += 1
            continue
        rm = report_start_re.match(line)
        if rm:
            store = rm.group(1).strip()
            block = {}
            i += 1
            if i < len(lines):
                tm = time_line_re.match(lines[i])
                if tm:
                    for nm in tm.group(1).split():
                        block[nm] = "정상출근"
                    i += 1
            while i < len(lines) and not end_re.match(lines[i]) and not report_start_re.match(lines[i]):
                parts = lines[i].split()
                if parts and parts[0] in _KAKAO_STATUS_MAP:
                    code = _KAKAO_STATUS_MAP[parts[0]]
                    for nm in parts[1:]:
                        block[nm] = code
                i += 1
            if i < len(lines) and end_re.match(lines[i]):
                i += 1
            for name, code in block.items():
                results.append({
                    "raw": f"{store} 출근보고", "name": name, "store": store,
                    "work_date": current_date, "code": code, "memo": "",
                })
            continue
        dm = _DATE_DIVIDER_RE.search(line)
        if dm:
            try:
                current_date = date(int(dm.group(1)), int(dm.group(2)), int(dm.group(3)))
            except ValueError:
                pass
        i += 1
    return results


def parse_sosajang_chat(text, fallback_date):
    """소사장 카톡 형식: "매장명 출근입니다" / "매장명 휴무입니다 매장은 CLOSE 입니다" 등 한 줄 자기 보고.
    "금일 휴무입니다"처럼 메시지에 매장명이 없으면, 바로 위 발신자 줄("대교_매장명 이름소사장")에서
    매장명을 찾아 이어서 사용."""
    msg_re = re.compile(r"^(.+?)\s*(출근|휴무|퇴근)입니다")
    short_stores = [s.replace("대교대리점 ", "") for s in db.SOSAJANG_STORES]
    no_store_words = {"금일", "오늘", "익일", "내일"}

    results = []
    current_date = fallback_date
    current_store = None
    for raw in text.splitlines():
        line = _strip_kakao_prefix(raw)
        if not line:
            continue

        store_matches = [s for s in short_stores if s in line]
        found_store = max(store_matches, key=len) if store_matches else None

        mm = msg_re.match(line)
        if mm:
            body_store = mm.group(1).strip()
            if found_store:
                store = found_store
            elif body_store in no_store_words:
                store = current_store or body_store
            else:
                store = body_store
            status = mm.group(2)
            code = "휴무" if status == "휴무" else "정상출근"
            memo = "CLOSE" if "CLOSE" in line.upper() else ""
            results.append({
                "raw": line, "name": None, "store": store,
                "work_date": current_date, "code": code, "memo": memo,
            })
            if found_store:
                current_store = found_store
            continue

        dm = _DATE_DIVIDER_RE.search(line)
        if dm:
            try:
                current_date = date(int(dm.group(1)), int(dm.group(2)), int(dm.group(3)))
            except ValueError:
                pass
            continue

        if found_store:
            current_store = found_store
    return results


_KAKAO_EXAMPLES = {
    "본사": "8월 28일 (금) 전산팀 출근현황\n최요안나 : 출근\n동희원 : 연차\n김선영 : 휴무",
    "직영": "상동점 출근보고\n10시 유경학\n연차 김찬양\n이상입니다",
    "소사장": "군포역점 출근입니다\n인덕원점 휴무입니다 매장은 CLOSE 입니다",
}


def _decode_uploaded_text(uploaded_file):
    raw_bytes = uploaded_file.getvalue()
    for enc in ("utf-8-sig", "utf-8", "cp949", "euc-kr"):
        try:
            return raw_bytes.decode(enc)
        except (UnicodeDecodeError, LookupError):
            continue
    return raw_bytes.decode("utf-8", errors="ignore")


def kakao_import_view(user):
    st.subheader("카톡 근태 가져오기")
    st.caption(
        "카톡방(본사/직영/소사장)에서 '대화 내보내기'로 받은 .txt 파일을 통째로 올리면, 그중 아래에서 고른 "
        "달(월)의 근태만 걸러서 정리해드려요. 확인하고 저장하면 '근태표 양식 엑셀 다운로드'에도 바로 반영돼요. "
        "직원들은 지금처럼 카톡에만 올리면 되고, 따로 앱에 입력할 필요 없어요."
    )

    category = st.selectbox("구분", db.CATEGORIES, key="kakao_category")
    c1, c2 = st.columns(2)
    with c1:
        target_year = st.number_input(
            "가져올 연도", min_value=2020, max_value=2100, value=date.today().year, step=1, key="kakao_year"
        )
    with c2:
        target_month = st.number_input(
            "가져올 월", min_value=1, max_value=12, value=date.today().month, step=1, key="kakao_month"
        )
    st.caption(f"→ {int(target_year)}년 {int(target_month)}월 근태만 걸러서 가져와요. 나머지 달 내용은 자동으로 무시됩니다.")

    uploaded = st.file_uploader(
        "카톡 대화 파일 업로드 (.txt, 카카오톡 '대화 내보내기')", type=["txt"], key="kakao_file"
    )
    chat_text = st.text_area(
        "또는 카톡 대화 내용 직접 붙여넣기", height=180,
        placeholder=f"예시:\n{_KAKAO_EXAMPLES[category]}", key="kakao_text",
    )

    if st.button("미리보기", key="kakao_preview_btn", use_container_width=True):
        source_text = _decode_uploaded_text(uploaded) if uploaded is not None else chat_text
        if not source_text or not source_text.strip():
            st.warning("파일을 올리거나 대화 내용을 붙여넣어주세요.")
            st.session_state.pop("kakao_parsed", None)
        else:
            fallback_date = date(int(target_year), int(target_month), 1)
            if category == "본사":
                all_parsed = parse_headquarters_chat(source_text, fallback_date)
            elif category == "직영":
                all_parsed = parse_jikyeong_chat(source_text, fallback_date)
            else:
                all_parsed = parse_sosajang_chat(source_text, fallback_date)

            parsed = [
                r for r in all_parsed
                if r["work_date"].year == int(target_year) and r["work_date"].month == int(target_month)
            ]
            st.session_state["kakao_parsed"] = parsed
            st.session_state["kakao_parsed_category"] = category
            st.info(f"전체 인식 {len(all_parsed)}건 중 {int(target_year)}년 {int(target_month)}월 대상 {len(parsed)}건만 가져왔어요.")
            if not parsed:
                st.warning("인식된 근태 내용이 없어요. 형식이 다르면 캡처해서 알려주시면 맞춰드릴게요.")

    parsed = st.session_state.get("kakao_parsed")
    parsed_category = st.session_state.get("kakao_parsed_category")
    if parsed and parsed_category == category:
        st.divider()
        st.markdown(f"**미리보기 — {len(parsed)}건 인식됨. 확인하고 틀린 부분은 고친 뒤 저장하세요.**")

        cat_employees = [e for e in db.list_users(include_inactive=False) if e["category"] == category]
        emp_names = [e["name"] for e in cat_employees]

        with st.form("kakao_save_form"):
            row_keys = []
            for idx, row in enumerate(parsed):
                with st.container(border=True):
                    st.caption(f"원문: {row['raw']}")
                    c1, c2, c3, c4, c5 = st.columns([1.3, 1.6, 1.3, 1.6, 0.8])
                    with c1:
                        st.date_input(
                            "날짜", value=row["work_date"], key=f"kk_date_{idx}", label_visibility="collapsed"
                        )
                    with c2:
                        if category == "소사장" and row.get("store"):
                            match = next(
                                (e for e in cat_employees if row["store"] in e["department"]), None
                            )
                        else:
                            match = next((e for e in cat_employees if e["name"] == row.get("name")), None)
                        options = ["(직접 선택)"] + emp_names
                        default_idx = options.index(match["name"]) if match and match["name"] in options else 0
                        st.selectbox(
                            "직원", options, index=default_idx, key=f"kk_emp_{idx}", label_visibility="collapsed"
                        )
                    with c3:
                        code_default = row["code"] if row["code"] in db.ATTENDANCE_CODES else "기타"
                        st.selectbox(
                            "근태코드", db.ATTENDANCE_CODES, index=db.ATTENDANCE_CODES.index(code_default),
                            key=f"kk_code_{idx}", label_visibility="collapsed",
                        )
                    with c4:
                        st.text_input(
                            "메모", value=row.get("memo", ""), key=f"kk_memo_{idx}", label_visibility="collapsed"
                        )
                    with c5:
                        st.checkbox("반영", value=True, key=f"kk_include_{idx}")
                    row_keys.append(idx)

            submitted = st.form_submit_button("선택한 내용 저장", use_container_width=True)

        if submitted:
            saved = 0
            unmatched = 0
            for idx in row_keys:
                if not st.session_state.get(f"kk_include_{idx}", True):
                    continue
                chosen_name = st.session_state.get(f"kk_emp_{idx}")
                if chosen_name == "(직접 선택)" or not chosen_name:
                    unmatched += 1
                    continue
                emp = next((e for e in cat_employees if e["name"] == chosen_name), None)
                if not emp:
                    unmatched += 1
                    continue
                wd = st.session_state.get(f"kk_date_{idx}")
                code = st.session_state.get(f"kk_code_{idx}")
                memo = st.session_state.get(f"kk_memo_{idx}", "")
                db.upsert_attendance(emp["id"], wd.isoformat(), code, memo)
                saved += 1
            if saved:
                st.success(f"{saved}건 저장되었습니다.")
            if unmatched:
                st.warning(f"직원이 매칭되지 않아 저장하지 않은 항목이 {unmatched}건 있어요. '직원' 항목에서 직접 선택 후 다시 저장해주세요.")
            else:
                del st.session_state["kakao_parsed"]
                st.rerun()


# ---------------------------------------------------------------------------
# 원본 근태 파일 형식(본사근태 / 직영근태 / 소사장근태) 그대로 엑셀 다운로드
# ---------------------------------------------------------------------------

# 본사/직영: 매장(부서) + 직원명 행 x 날짜 열 그리드. 예전 본사근태/직영근태 파일과 동일한 구조.
def _build_roster_matrix_sheet(wb, sheet_name, employees, att_map, dates):
    ws = wb.create_sheet(title=sheet_name)
    weekday_kr = ["월", "화", "수", "목", "금", "토", "일"]
    n = len(dates)

    ws.cell(row=1, column=1, value=f"{sheet_name} 일정표 (빈 칸 = 정상출근)")
    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=6 + n)

    ws.cell(row=2, column=1, value="부서/매장")
    ws.cell(row=2, column=2, value="직원명")
    for i, d in enumerate(dates):
        ws.cell(row=2, column=3 + i, value=weekday_kr[d.weekday()])
    ws.cell(row=2, column=3 + n, value="휴무")
    ws.cell(row=2, column=4 + n, value="연차")
    ws.cell(row=2, column=5 + n, value="반차")
    ws.cell(row=2, column=6 + n, value="합계")

    for i, d in enumerate(dates):
        ws.cell(row=3, column=3 + i, value=d.strftime("%-m/%-d"))

    r = 4
    for e in sorted(employees, key=lambda x: (x["department"], x["name"])):
        ws.cell(row=r, column=1, value=e["department"])
        ws.cell(row=r, column=2, value=e["name"])
        hol = annual = half = 0
        for i, d in enumerate(dates):
            code = att_map.get((e["id"], d.isoformat()))
            ws.cell(row=r, column=3 + i, value="" if code in (None, "정상출근") else code)
            if code == "휴무":
                hol += 1
            elif code == "연차":
                annual += 1
            elif code == "반차":
                half += 1
        ws.cell(row=r, column=3 + n, value=hol)
        ws.cell(row=r, column=4 + n, value=annual)
        ws.cell(row=r, column=5 + n, value=half)
        ws.cell(row=r, column=6 + n, value=hol + annual + half)
        r += 1

    ws.cell(row=r, column=1, value="재직인원")
    for i in range(n):
        ws.cell(row=r, column=3 + i, value=len(employees))


# 소사장: 개인별 근태코드를 예전 '출첵' 파일처럼 매장 단위 요약(o/휴/개인/조기퇴근/지각/미입력)으로 변환
_STORE_STATUS_MAP = {
    "정상출근": "o",
    "특근": "o",
    "당직": "o",
    "교육": "o",
    "휴무": "휴",
    "연차": "개인",
    "반차": "개인",
    "무급": "개인",
    "개인용무": "개인",
    "경조": "개인",
    "예비군": "개인",
    "지각": "지각",
    "조퇴": "조기퇴근",
    "기타": "기타",
}


def _store_status_for_date(codes):
    """같은 매장 소속 직원들의 그날 근태코드 목록 -> 매장 단위 상태 문자열로 요약"""
    if not codes:
        return "미입력"
    statuses = [_STORE_STATUS_MAP.get(c, c) for c in codes]
    if "o" in statuses:
        return "o"
    if all(s == "휴" for s in statuses):
        return "휴"
    uniq = []
    for s in statuses:
        if s not in uniq:
            uniq.append(s)
    return "/".join(uniq)


def _build_store_summary_sheet(wb, employees, att_map, dates):
    ws = wb.create_sheet(title="소사장근태")
    weekday_kr = ["월", "화", "수", "목", "금", "토", "일"]
    n = len(dates)

    seen = set()
    stores = []
    for e in employees:
        if e["department"] not in seen:
            stores.append(e["department"])
            seen.add(e["department"])
    ordered_stores = [s for s in db.SOSAJANG_STORES if s in seen] + [s for s in stores if s not in db.SOSAJANG_STORES]

    ws.cell(
        row=1, column=1,
        value="소사장 매장별 근태 요약 (o=정상, 휴=휴무, 개인=연차·반차 등 개인사유, 조기퇴근, 지각, 미입력=입력 없음)",
    )
    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=1 + n)
    ws.cell(row=2, column=1, value="매장명")
    for i, d in enumerate(dates):
        ws.cell(row=2, column=2 + i, value=weekday_kr[d.weekday()])
    for i, d in enumerate(dates):
        ws.cell(row=3, column=2 + i, value=d.strftime("%-m/%-d"))

    emp_by_store = {}
    for e in employees:
        emp_by_store.setdefault(e["department"], []).append(e)

    r = 4
    for store in ordered_stores:
        ws.cell(row=r, column=1, value=store)
        for i, d in enumerate(dates):
            ds = d.isoformat()
            codes = [att_map.get((e["id"], ds)) for e in emp_by_store.get(store, [])]
            codes = [c for c in codes if c]
            ws.cell(row=r, column=2 + i, value=_store_status_for_date(codes))
        r += 1


def build_category_format_excel(start_date, end_date):
    """본사근태 / 직영근태 / 소사장근태 - 예전에 쓰던 파일과 같은 서식으로 엑셀 생성 (구분별 시트 분리)"""
    dates = []
    d = start_date
    while d <= end_date:
        dates.append(d)
        d += timedelta(days=1)
    if not dates:
        return None

    all_employees = db.list_users(include_inactive=False)
    wb = Workbook()
    wb.remove(wb.active)
    made_sheet = False

    for category, sheet_name in [("본사", "본사근태"), ("직영", "직영근태")]:
        emps = [e for e in all_employees if e["category"] == category]
        if not emps:
            continue
        records = db.get_attendance_records_for_matrix(category, start_date.isoformat(), end_date.isoformat())
        att_map = {(r["user_id"], r["work_date"]): r["code"] for r in records}
        _build_roster_matrix_sheet(wb, sheet_name, emps, att_map, dates)
        made_sheet = True

    sosajang_emps = [e for e in all_employees if e["category"] == "소사장"]
    if sosajang_emps:
        records = db.get_attendance_records_for_matrix("소사장", start_date.isoformat(), end_date.isoformat())
        att_map = {(r["user_id"], r["work_date"]): r["code"] for r in records}
        _build_store_summary_sheet(wb, sosajang_emps, att_map, dates)
        made_sheet = True

    if not made_sheet:
        return None
    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()


# ---------------------------------------------------------------------------
# 관리자용: 전체 현황 + 계정 관리
# ---------------------------------------------------------------------------
def admin_view(user):
    tab1, tab_kakao, tab2 = st.tabs(["전체 근태 현황", "카톡 근태 가져오기", "직원 계정 관리"])

    with tab_kakao:
        kakao_import_view(user)

    with tab1:
        st.subheader("근태 미입력 매장/부서 확인")
        check_date = st.date_input("확인할 날짜", value=date.today(), key="missing_check_date")
        summary = db.get_org_unit_summary_for_date(check_date.isoformat())
        missing = [s for s in summary if s["employee_count"] > 0 and s["entered_count"] == 0]
        if missing:
            st.warning(f"{check_date.isoformat()} 기준, 근태 입력이 하나도 없는 부서/매장이 {len(missing)}곳 있어요.")
            miss_df = pd.DataFrame(missing)[["category", "department", "employee_count"]]
            miss_df.columns = ["구분", "부서/매장", "재직 인원"]
            st.dataframe(miss_df, use_container_width=True, hide_index=True)
        else:
            st.success(f"{check_date.isoformat()} 기준, 모든 부서/매장에 근태가 입력되었어요.")

        st.divider()
        st.subheader("전체 근태 현황")
        c1, c2, c3, c4 = st.columns(4)
        with c1:
            start = st.date_input(
                "조회 시작일", value=date.today().replace(day=1), key="admin_start"
            )
        with c2:
            end = st.date_input("조회 종료일", value=date.today(), key="admin_end")
        with c3:
            category = st.selectbox("구분", ["전체"] + db.CATEGORIES, key="admin_category")
        with c4:
            if category == "전체":
                org_options = ["전체"] + db.DEPARTMENTS + db.JIKYEONG_STORES + db.SOSAJANG_STORES
            else:
                org_options = ["전체"] + db.get_org_units(category)
            org_unit = st.selectbox("부서/매장", org_options, key=f"admin_org_{category}")

        st.markdown("**근태표 양식으로 다운로드 (본사·직영·소사장 원본 서식)**")
        st.caption(
            "예전에 쓰시던 본사근태·직영근태 파일과 같은 매장/부서·직원명 × 날짜 표, "
            "소사장근태(출첵) 파일과 같은 매장 단위 요약표로 시트를 나눠서 받아요. "
            "위 조회 시작일/종료일 기준으로 만들어집니다."
        )
        matrix_bytes = build_category_format_excel(start, end)
        if matrix_bytes:
            st.download_button(
                "근태표 양식 엑셀 다운로드",
                data=matrix_bytes,
                file_name=f"근태표_{start.isoformat()}_{end.isoformat()}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                key="matrix_download",
            )
        else:
            st.info("등록된 직원이 없어 근태표를 만들 수 없습니다.")

        st.divider()

        records = db.get_all_attendance(start.isoformat(), end.isoformat(), category, org_unit)
        if records:
            df = pd.DataFrame(records)[
                ["work_date", "category", "department", "name", "code", "memo", "updated_at"]
            ]
            df.columns = ["날짜", "구분", "부서/매장", "이름", "근태코드", "메모", "최종수정"]
            st.dataframe(df, use_container_width=True, hide_index=True)

            # 엑셀 다운로드
            buf = io.BytesIO()
            with pd.ExcelWriter(buf, engine="openpyxl") as writer:
                df.to_excel(writer, index=False, sheet_name="근태현황")
            st.download_button(
                "엑셀로 다운로드",
                data=buf.getvalue(),
                file_name=f"근태현황_{start.isoformat()}_{end.isoformat()}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )
        else:
            st.info("해당 조건에 입력된 근태 내역이 없습니다.")

    with tab2:
        st.subheader("직원 계정 추가")

        c1, c2 = st.columns(2)
        with c1:
            add_category = st.selectbox("구분", db.CATEGORIES, key="add_user_category")
        with c2:
            add_org_options = db.get_org_units(add_category)
            add_org_unit = st.selectbox(
                "부서/매장", add_org_options, key=f"add_user_org_{add_category}"
            )

        with st.form("add_user_form"):
            c1, c2 = st.columns(2)
            with c1:
                new_username = st.text_input("아이디 (사번 등, 영문/숫자 권장)")
                new_name = st.text_input("이름")
            with c2:
                new_password = st.text_input("초기 비밀번호", type="password")
            new_role = st.radio("권한", ["employee", "admin"], horizontal=True,
                                 format_func=lambda x: "일반 직원" if x == "employee" else "관리자")
            submitted = st.form_submit_button("계정 생성", use_container_width=True)
        if submitted:
            if not new_username or not new_password or not new_name:
                st.error("아이디, 이름, 초기 비밀번호는 필수입니다.")
            else:
                ok, msg = db.create_user(
                    new_username.strip(), new_password, new_name.strip(),
                    add_category, add_org_unit, new_role,
                )
                if ok:
                    st.success(f"{msg} ({add_category} · {add_org_unit})")
                    st.rerun()
                else:
                    st.error(msg)

        st.divider()
        with st.expander("엑셀로 직원 일괄 등록 (여러 명 한번에)"):
            st.caption(
                "70명을 한 명씩 등록하기 번거로우시면, 아래 양식에 맞춰 엑셀을 채워서 올리세요. "
                "'구분'은 본사/직영/소사장 중 하나, '부서/매장'은 그 구분에 실제로 있는 이름과 정확히 같아야 해요."
            )
            st.download_button(
                "빈 양식 다운로드 (예시 3줄 포함)",
                data=_build_bulk_template(),
                file_name="직원_일괄등록_양식.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                key="bulk_template_dl",
            )

            uploaded = st.file_uploader(
                "작성한 엑셀 파일 업로드 (.xlsx)", type=["xlsx"], key="bulk_upload_file"
            )
            if uploaded is not None:
                try:
                    in_df = pd.read_excel(uploaded, dtype=str).fillna("")
                except Exception as e:
                    st.error(f"파일을 읽을 수 없습니다: {e}")
                    in_df = None

                if in_df is not None:
                    required_cols = {"아이디", "이름", "구분", "부서/매장"}
                    if not required_cols.issubset(set(in_df.columns)):
                        st.error(f"필수 열이 빠졌습니다. 필요한 열: {', '.join(sorted(required_cols))}")
                    else:
                        st.dataframe(in_df, use_container_width=True, hide_index=True)
                        if st.button("일괄 등록 실행", key="bulk_register_btn", use_container_width=True):
                            success, failed = [], []
                            for _, row in in_df.iterrows():
                                uname = str(row.get("아이디", "")).strip()
                                name = str(row.get("이름", "")).strip()
                                category = str(row.get("구분", "")).strip()
                                org_unit = str(row.get("부서/매장", "")).strip()
                                pw = str(row.get("초기비밀번호", "")).strip() or "changeme123"
                                role = str(row.get("권한", "")).strip()
                                role = role if role in ("employee", "admin") else "employee"

                                if not uname or not name:
                                    failed.append((uname or "(빈값)", "아이디 또는 이름이 비어있음"))
                                    continue
                                if category not in db.CATEGORIES:
                                    failed.append((uname, f"구분 '{category}'이 본사/직영/소사장 중 하나가 아님"))
                                    continue
                                if org_unit not in db.get_org_units(category):
                                    failed.append((uname, f"'{category}'에 '{org_unit}' 부서/매장이 없음"))
                                    continue

                                ok, msg = db.create_user(uname, pw, name, category, org_unit, role)
                                (success if ok else failed).append(uname if ok else (uname, msg))

                            st.success(f"성공: {len(success)}건")
                            if failed:
                                st.error(f"실패: {len(failed)}건")
                                st.dataframe(
                                    pd.DataFrame(failed, columns=["아이디", "실패 사유"]),
                                    use_container_width=True, hide_index=True,
                                )
                            if success:
                                st.rerun()

        st.divider()
        st.subheader("직원 목록")
        users = db.list_users()
        udf = pd.DataFrame(users)
        if not udf.empty:
            udf_display = udf[["username", "name", "category", "department", "role"]].copy()
            udf_display.columns = ["아이디", "이름", "구분", "부서/매장", "권한"]
            st.dataframe(udf_display, use_container_width=True, hide_index=True)

            with st.expander("복사하기 편한 텍스트로 보기"):
                st.caption("이 표는 그림처럼 그려져서 Ctrl+C가 잘 안 먹혀요. 아래 상자 안 글자는 평범하게 드래그해서 복사하시면 돼요.")
                lines = [
                    f"{u['username']}\t{u['name']}\t{u['category']}\t{u['department']}\t{u['role']}"
                    for u in users
                ]
                st.text_area(
                    "아이디 / 이름 / 구분 / 부서·매장 / 권한",
                    value="아이디\t이름\t구분\t부서/매장\t권한\n" + "\n".join(lines),
                    height=200, key="user_list_copy_text",
                )

            with st.expander("직원 정보 수정 / 계정 삭제"):
                edit_target = st.selectbox(
                    "대상 직원",
                    users,
                    format_func=lambda u: f"{u['name']} ({u['username']}, {u['category']} · {u['department']})",
                    key="edit_target_select",
                )

                e_c1, e_c2 = st.columns(2)
                with e_c1:
                    edit_name = st.text_input(
                        "이름", value=edit_target["name"], key=f"edit_name_{edit_target['id']}"
                    )
                    edit_category = st.selectbox(
                        "구분", db.CATEGORIES,
                        index=db.CATEGORIES.index(edit_target["category"])
                        if edit_target["category"] in db.CATEGORIES else 0,
                        key=f"edit_category_{edit_target['id']}",
                    )
                with e_c2:
                    edit_org_options = db.get_org_units(edit_category)
                    edit_department = st.selectbox(
                        "부서/매장", edit_org_options,
                        index=edit_org_options.index(edit_target["department"])
                        if edit_target["department"] in edit_org_options else 0,
                        key=f"edit_dept_{edit_target['id']}_{edit_category}",
                    )
                    edit_role = st.radio(
                        "권한", ["employee", "admin"], horizontal=True,
                        index=0 if edit_target["role"] == "employee" else 1,
                        format_func=lambda x: "일반 직원" if x == "employee" else "관리자",
                        key=f"edit_role_{edit_target['id']}",
                    )
                edit_new_pw = st.text_input(
                    "새 비밀번호 (바꿀 때만 입력, 비워두면 기존 비밀번호 유지)",
                    type="password", key=f"edit_pw_{edit_target['id']}",
                )

                if st.button("수정 저장", key=f"edit_save_{edit_target['id']}", use_container_width=True):
                    db.update_user(
                        edit_target["id"], edit_name.strip(), edit_category, edit_department,
                        edit_role, edit_new_pw or None,
                    )
                    st.success(f"{edit_name} 계정 정보가 수정되었습니다.")
                    st.rerun()

                st.divider()
                confirm_del = st.checkbox(
                    "정말 삭제할게요 (근태 입력 내역도 함께 삭제되며, 되돌릴 수 없어요)",
                    key=f"confirm_del_{edit_target['id']}",
                )
                if st.button(
                    "계정 완전 삭제", key=f"edit_del_{edit_target['id']}",
                    type="secondary", use_container_width=True, disabled=not confirm_del,
                ):
                    if edit_target["username"] == "admin":
                        st.error("기본 admin 계정은 삭제할 수 없습니다.")
                    else:
                        db.delete_user(edit_target["id"])
                        st.success(f"{edit_target['name']} 계정이 완전히 삭제되었습니다.")
                        st.rerun()

            with st.expander("계정 비활성화(퇴사 처리)"):
                target = st.selectbox(
                    "비활성화할 직원",
                    users,
                    format_func=lambda u: f"{u['name']} ({u['username']}, {u['category']} · {u['department']})",
                    key="deactivate_target_select",
                )
                if st.button("선택한 계정 비활성화", type="secondary", key="deactivate_btn"):
                    if target["username"] == "admin":
                        st.error("기본 admin 계정은 비활성화할 수 없습니다.")
                    else:
                        db.deactivate_user(target["id"])
                        st.success(f"{target['name']} 계정이 비활성화되었습니다.")
                        st.rerun()

            with st.expander("비활성화된 계정 되살리기"):
                inactive_users = [u for u in db.list_users(include_inactive=True) if u["active"] == 0]
                if not inactive_users:
                    st.caption("비활성화된 계정이 없습니다.")
                else:
                    revive_target = st.selectbox(
                        "되살릴 직원",
                        inactive_users,
                        format_func=lambda u: f"{u['name']} ({u['username']}, {u['category']} · {u['department']})",
                        key="revive_target_select",
                    )
                    if st.button("선택한 계정 되살리기", key="revive_btn", use_container_width=True):
                        db.activate_user(revive_target["id"])
                        st.success(f"{revive_target['name']} 계정이 다시 활성화되었습니다.")
                        st.rerun()


# ---------------------------------------------------------------------------
# 메인
# ---------------------------------------------------------------------------
def main():
    if "user" not in st.session_state:
        login_view()
        return

    user = st.session_state["user"]
    logout_button()

    st.title("🗂️ (주)대교통신 HR 플랫폼")

    if user["role"] == "admin":
        tab_emp, tab_bulk, tab_notice, tab_docs, tab_onboard, tab_admin = st.tabs(
            ["내 근태입력", "일괄입력", "공지사항", "문서함", "온보딩", "관리자"]
        )
        with tab_emp:
            employee_view(user)
        with tab_bulk:
            bulk_entry_view(user)
        with tab_notice:
            announcements_view(user)
        with tab_docs:
            documents_view("규정", "아직 등록된 규정 문서가 없습니다. GitHub 저장소의 docs/규정 폴더에 파일을 올려주세요.")
        with tab_onboard:
            documents_view("온보딩", "아직 등록된 온보딩 자료가 없습니다. GitHub 저장소의 docs/온보딩 폴더에 파일을 올려주세요.")
        with tab_admin:
            admin_view(user)
    else:
        tab_emp, tab_bulk, tab_notice, tab_docs, tab_onboard = st.tabs(
            ["내 근태입력", "일괄입력", "공지사항", "문서함", "온보딩"]
        )
        with tab_emp:
            employee_view(user)
        with tab_bulk:
            bulk_entry_view(user)
        with tab_notice:
            announcements_view(user)
        with tab_docs:
            documents_view("규정", "아직 등록된 규정 문서가 없습니다.")
        with tab_onboard:
            documents_view("온보딩", "아직 등록된 온보딩 자료가 없습니다.")


if __name__ == "__main__":
    main()
