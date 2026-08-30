"""
(주)대교통신 사내 HR 플랫폼 - Phase 1 (근태입력 셀프서비스 MVP)

실행:
    streamlit run app.py

기본 관리자 계정: admin / changeme123  (반드시 최초 로그인 후 비밀번호를 바꿀 직원용 계정을 새로 만들고,
                                       admin 계정 비밀번호도 바꿔서 사용하세요)
"""

import io
from datetime import date, timedelta

import pandas as pd
import streamlit as st

import db

st.set_page_config(page_title="(주)대교통신 HR 플랫폼", page_icon="🗂️", layout="wide")
db.init_db()


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

    with st.form("bulk_attendance_form"):
        entries = {}
        for m in members:
            prev = existing.get(m["id"], {})
            prev_code = prev.get("code", "정상출근")
            prev_memo = prev.get("memo", "")
            code_idx = db.ATTENDANCE_CODES.index(prev_code) if prev_code in db.ATTENDANCE_CODES else 0

            col_name, col_code, col_memo = st.columns([2, 2, 3])
            with col_name:
                st.markdown(f"**{m['name']}**")
            with col_code:
                code = st.selectbox(
                    "근태 코드", db.ATTENDANCE_CODES, index=code_idx,
                    key=f"bulk_code_{m['id']}", label_visibility="collapsed",
                )
            with col_memo:
                memo = st.text_input(
                    "메모", value=prev_memo, key=f"bulk_memo_{m['id']}",
                    label_visibility="collapsed", placeholder="메모 (선택)",
                )
            entries[m["id"]] = (code, memo)

        submitted = st.form_submit_button(f"{org_unit} 전체 저장 ({work_date.isoformat()})", use_container_width=True)

    if submitted:
        for user_id, (code, memo) in entries.items():
            db.upsert_attendance(user_id, work_date.isoformat(), code, memo)
        st.success(f"{work_date.isoformat()} 기준 {org_unit} {len(entries)}명 근태가 저장되었습니다.")
        st.rerun()


# ---------------------------------------------------------------------------
# 관리자용: 전체 현황 + 계정 관리
# ---------------------------------------------------------------------------
def admin_view(user):
    tab1, tab2 = st.tabs(["전체 근태 현황", "직원 계정 관리"])

    with tab1:
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
        st.subheader("직원 목록")
        users = db.list_users()
        udf = pd.DataFrame(users)
        if not udf.empty:
            udf_display = udf[["username", "name", "category", "department", "role"]].copy()
            udf_display.columns = ["아이디", "이름", "구분", "부서/매장", "권한"]
            st.dataframe(udf_display, use_container_width=True, hide_index=True)

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
        tab_emp, tab_bulk, tab_admin = st.tabs(["내 근태입력", "일괄입력", "관리자"])
        with tab_emp:
            employee_view(user)
        with tab_bulk:
            bulk_entry_view(user)
        with tab_admin:
            admin_view(user)
    else:
        tab_emp, tab_bulk = st.tabs(["내 근태입력", "일괄입력"])
        with tab_emp:
            employee_view(user)
        with tab_bulk:
            bulk_entry_view(user)


if __name__ == "__main__":
    main()
