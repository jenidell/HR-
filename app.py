"""
대교통신 사내 HR 플랫폼 - Phase 1 (근태입력 셀프서비스 MVP)

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

st.set_page_config(page_title="대교통신 HR 플랫폼", page_icon="🗂️", layout="wide")
db.init_db()


# ---------------------------------------------------------------------------
# 로그인 / 세션 관리
# ---------------------------------------------------------------------------
def login_view():
    st.title("🗂️ 대교통신 HR 플랫폼")
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
        st.markdown(f"**{user['name']}** ({user['department']})")
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
 db.upert_attendance(user["id"), work_date.isoformat (), 코드, 메모)
 st.success(f"{work_date.isoformat ()} 근태가 '{code}'(으) 로 저장되었습니다)")
 세인트 rerun()

 col2와 함께:
 세인트 마크다운 ("**근태 코드 안내**)
 세인트 caption(
 "정상출근 · 연차 · 반차 · 지각 · 조퇴 · 특근 · 당직 · 교육 · "
 "경조 · 예비군 · 무급 · 개인용무 · 휴무 · 기타\n\n"
 "같은 날짜에 다시 입력하면 기존 내용을 덮어씁니다(수정)."
 )

 세인트디바이더 ()
 세인트 서브헤더 ("내 최근 입력 내역")

 default_start = 날짜.오늘 () - 시간 델타(일수=30)
 c1, c2 = st.columns(2)
 c1과 함께:
 시작 = st.date_input ("조회 시작일", 값=default_start, key="emp_start")
 c2와 함께:
 끝 = st.date_input ("조회 종료일", 값 = 날짜.오늘 (), 키="emp_end")

 레코드 = db.get_user_attendance(user["id"), start.isoformat (), end.isoformat ())
 기록된 경우:
 df = pd.데이터프레임(기록)[["work_date", "code", "memo", "updated_at"]
 df.columns = ["날짜", "근태코드", "메모", "최종수정"]
 st.dataframe(df, use_container_width=True, hide_index=True)
 그렇지 않으면:
 st.info("해당 기간에 입력된 근태 내역이 없습니다.")


# ---------------------------------------------------------------------------
# 관리자용: 전체 현황 + 계정 관리
# ---------------------------------------------------------------------------
def admin_view(사용자):
 tab1, tab2 = st.tabs(["전체 근태 현황", "직원 계정 관리"])

 탭1:
 세인트 서브헤더 ("전체 근태 현황")
 c1, c2, c3 = st.columns(3)
 c1과 함께:
 시작 = st.date_input(
 "조회 시작일", 값=날짜.오늘 ().replace(day=1), key="admin_start"
 )
 c2와 함께:
 끝 = st.date_input ("조회 종료일", 값 = 날짜.오늘 (), 키="admin_end")
 c3와 함께:
 = st.selectbox ("부서, ["전체"] + db.부서, key="admin_dep")

 레코드 = db.get_all_attendance (start.isoformat (), end.isoformat (), dep)
 기록된 경우:
 df = pd.데이터프레임(기록)[
 ["work_date", "부서", "이름", "코드", "memo", "updated_at"]
 ]
 df.columns = ["날짜", "부서", "이름", "근태코드", "메모", "최종수정"]
 st.dataframe(df, use_container_width=True, hide_index=True)

 # 엑셀 다운로드
 buf = IO.바이트IO()
 PD와 함께.ExcelWriter(buf, engine="openpyxl") 작성자:
 df.to _excel(작가, 인덱스=false, 시트_name="근태현황")
 st. download_버튼(
 "엑셀로 다운로드",
 data=buf.getvalue (),
 file_name=f"근태현황_{start.isoformat ()}_{end.isoformat ()}.xlsx",
 mime="application/vnd.openxmlformats-officeddocument.spreadsheetml.".시트",
 )
 그렇지 않으면:
 st.info("해당 조건에 입력된 근태 내역이 없습니다.")

 탭2:
 세인트 서브헤더 ("직원 계정 추가")
 st.form ("add_user_form 포함):
 c1, c2 = st.columns(2)
 c1과 함께:
 new_username = st.text_input ("아이디 (사번 등, 영문/숫자 권장))"
 new_name = st.text_input ("이름")
 c2와 함께:
 new_password = st.text_input ("초기 비밀번호", type="password")
 new_dep = st.selectbox ("부서", db.부서)
 new_role = st.radio ("권한, ["employee", "admin"], 가로=맞아요,
 format_func=lambda x: x == "employee"이 아닌 경우 "일반 직원")
 제출된 = st.form_submit_button ("계정 생성", use_container_width=True)
 제출된 경우:
 new_username 또는 new_password 또는 new_name이 아닌 경우:
 st.error("아이디, 이름, 초기 비밀번호는 필수입니다.")
 그렇지 않으면:
 알겠습니다, msg = db.create_user(
 new_username.strip (), new_password, new_name.strip (), new_dep, new_role
 )
 괜찮으시다면:
 st.success(msg)
 세인트 rerun()
 그렇지 않으면:
 st.error(msg)

 세인트디바이더 ()
 세인트 서브헤더 ("직원 목록")
 사용자 = db.list_users ()
 udf = pd.데이터프레임(사용자)
 udf.empty가 아니라면:
 udf_display = udf [[ "username", "이름", "부서", "역할"].copy ()
 udf_display.columns = ["아이디", "이름", "부서", "권한"]
 st.dataframe(udf_display, use_container_width=True, hide_index=True)

 with st.expander("계정 비활성화(퇴사 처리)"):
 대상 = st.selectbox(
 "비활성화할 직원",
 사용자,
 format_func=lambda u: f"{u['이름]}({u['username'], {u['부서]})",
 )
 st.button ("선택한 계정 비활성화", type="secondary":
 target["username"] == "admin"인 경우:
 st.error("기본 admin 계정은 비활성화할 수 없습니다.")
 그렇지 않으면:
 db.deactivate_user(target["id"])
 st.success(f"{target[이름]} 계정이 비활성화되었습니다.")
 세인트 rerun()


# ---------------------------------------------------------------------------
# 메인
# ---------------------------------------------------------------------------
주 () 정의:
 "사용자"가 st.session_state에 있지 않은 경우:
 login_view ()
 돌아가다

 사용자 = st.session_state["사용자"]
 로그아웃_버튼()

 st.title("🗂️ 대교통신 HR 플랫폼")

 사용자 ["역할"] == "관리자"인 경우:
 tab_emp, tab_admin = st. tabs(["내 근태입력", "관리자"])
 tab_emp와 함께:
 직원_view(사용자)
 tab_admin과 함께:
 admin_view(사용자)
 그렇지 않으면:
 직원_view(사용자)


__name__ == "__main__"인 경우:
 주된()
