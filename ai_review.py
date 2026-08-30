"""
ai_review.py
카톡 대화 내용을 Claude API로 검토하는 모듈.

하는 일:
  1) 카톡 대화에서 날짜별 근태 내용을 뽑아냄
  2) 오타 / 누락 / 중복 / 규정 위반 / 애매해서 사람이 판단해야 하는 건을 골라냄
  3) 결과를 정해진 JSON 구조로 돌려줌  →  app.py가 화면에 보여주고, 사람이 확인 후 DB에 반영

API 키는 코드나 GitHub에 절대 넣지 않습니다.
Streamlit Community Cloud → 앱 설정(Settings) → Secrets 에 아래 한 줄만 넣으면 됩니다.
    ANTHROPIC_API_KEY = "sk-ant-..."
(로컬에서 돌릴 때는 환경변수 ANTHROPIC_API_KEY 로도 동작합니다.)
"""

import json
import os
import re

# 기본 모델. 관리자 화면에서 바꿀 수 있습니다.
DEFAULT_MODEL = "claude-sonnet-4-5"

# 카톡 대화가 아주 길면 나눠서 검토합니다 (한 번에 보내는 글자 수 상한)
MAX_CHARS_PER_CHUNK = 55000


# ---------------------------------------------------------------------------
# API 키 / 클라이언트
# ---------------------------------------------------------------------------

def get_api_key():
    """Streamlit Secrets → 환경변수 순서로 API 키를 찾습니다."""
    try:
        import streamlit as st
        try:
            if "ANTHROPIC_API_KEY" in st.secrets:
                return str(st.secrets["ANTHROPIC_API_KEY"]).strip()
        except Exception:
            pass
    except Exception:
        pass
    key = os.environ.get("ANTHROPIC_API_KEY")
    return key.strip() if key else None


def is_configured():
    return bool(get_api_key())


def _client():
    try:
        import anthropic
    except ImportError as e:
        raise RuntimeError(
            "anthropic 패키지가 설치되어 있지 않습니다. "
            "requirements.txt에 anthropic 을 추가하고 다시 배포해주세요."
        ) from e
    key = get_api_key()
    if not key:
        raise RuntimeError(
            "ANTHROPIC_API_KEY가 설정되어 있지 않습니다.\n"
            "Streamlit Cloud → 앱 우측 상단 ⋮ → Settings → Secrets 에 다음 한 줄을 넣어주세요.\n"
            '    ANTHROPIC_API_KEY = "sk-ant-..."'
        )
    return anthropic.Anthropic(api_key=key)


def list_models():
    """이 계정에서 쓸 수 있는 모델 목록. 실패하면 빈 리스트."""
    try:
        resp = _client().models.list(limit=30)
        return [m.id for m in resp.data]
    except Exception:
        return []


# ---------------------------------------------------------------------------
# 구분별 판독 규정 (사용자가 정리해 둔 근태 규정 문서를 그대로 옮긴 것)
# ---------------------------------------------------------------------------

RULES_COMMON = """
[공통 원칙]
- 카톡에서 근거를 찾을 수 없는 날짜/사람은 절대 임의로 추측해서 채우지 않는다.
  근거가 없으면 records에 넣지 말고, 필요하면 issues에 '누락'으로 알린다.
- 규정에 없는 새로운 패턴이나 애매한 표현이 나오면 혼자 판단하지 말고
  issues에 kind='확인필요'로 넣어서 사람에게 물어본다.
- 오타로 보이는 이름/매장명/키워드는 issues에 kind='오타'로 넣고,
  suggestion에 '무엇으로 고치면 될지'를 적는다. 확신이 높으면 records에도 고친 값으로 넣되
  source에 '오타 추정 보정'이라고 적는다.
- 같은 날짜·같은 대상에 서로 다른 내용이 두 번 올라오면 kind='중복'으로 알린다.
"""

RULES_BONSA = """
[본사 근태 판독 규정]
- 대상은 '사람'(직원 이름)이다.
- 키워드 → 근태코드 매핑:
    휴무 → 휴무 / 연차 → 연차 / 반차 → 반차 / 예비군 → 예비군 / 지각 → 지각
    경조 → 경조 / 교육 → 교육 / 특근 → 특근 / 당직 → 당직 / 무급 → 무급
    (그 외 정상 출근은 → 정상출근)
- 반차와 무급은 0.5일로 계산되는 항목이다.
- 팀 구분: 관리팀 / 소매팀 / 도매팀 / 재고팀 / 전산팀
"""

RULES_JIKYEONG = """
[직영점 근태 판독 규정]
- 대상은 '사람'(직원 이름)이다.
- 데이터 우선순위:
    1) 출근보고 메시지를 1순위 근거로 쓴다.
       (예: '범계역직영점 출근보고 / 10시 홍길동 김철수 / 휴무 이영희')
    2) 어떤 매장의 특정 날짜에 출근보고가 아예 없고 퇴근보고만 있을 때에만 퇴근보고를 쓴다.
    3) 출근보고와 퇴근보고 내용이 다른 날이 자주 있다(직원이 저녁에 잘못 적는 경우).
       둘 다 있으면 항상 출근보고를 쓰고 퇴근보고는 무시한다.
       단, 둘의 내용이 다르면 issues에 kind='확인필요'로 같이 알려준다.
    4) 제목이 '퇴근보고'여도 전송 시각이 오전(출근 시간대)이면 실제로는 출근보고로 본다.
    5) 하안점·덕천직영점은 일요일이 정기휴무라 보고가 안 올라온다. 이건 누락으로 보지 않는다.
- 키워드 → 상태 매핑:
    (키워드 없이 이름만, 시간 표기와 함께) → 출
    휴무 → 휴무 / 연차 → 연차 / 지각 → 지각 / 예비군 → 예비군 / 특근 → 특근
    본사, 회의 → 본사회의 / 교육 → 삼성교육본사회의실
    덕천지원 → 덕천지원 / 하안지원 → 하안지원
    '지원'만 단독으로 쓰였으면 → '[해당 출근보고의 매장명]지원'
    이름은 나왔는데 보고한 매장이 그 직원의 소속 매장과 다르면 → '[보고 매장명]지원'
- 직영 매장: 범계역직영점 / 상동점 / 의왕점 / 덕천직영점 / 하안점
"""

RULES_SOSAJANG = """
[소사장 근태 판독 규정]
- 대상은 '사람'이 아니라 '매장'이다. target에는 반드시 매장명을 넣는다.
- 하루에 두 가지를 판단한다:
    code       = 출근보고 상태 (소사장 근태 양식 '출첵'의 하루 2칸 중 왼쪽)
    close_code = 퇴근보고 상태 (오른쪽 칸). 정상 마감이면 'o'.
                 휴무인 날은 마감보고를 안 하므로 close_code를 비워둔다.
- 코드 체계:
    o        정상 출근/마감
    휴       휴무
    미       매장 미오픈 (휴무 + 매장미오픈 둘 다 해당)
    개인     개인 용무
    휴가     휴가
    본사     본사 / 본사회의 (정보성 메모, 결근 아님)
    조기마감  월 1회 가능
    조기퇴근  월 1회 가능
- 코드가 2개 겹치면 '개인+조기마감'처럼 code에 슬래시로 결합해서 넣는다. (예: '개인/조기마감')
- '대표님승인건', '가족상', '병원진료' 같은 예외 문구는 원문 그대로 memo에 넣는다.
- 규정에 없는 새 패턴이 나오면 절대 임의 판단하지 말고 반드시 issues에 kind='확인필요'로 넣는다.
- 매장명은 '대교대리점 XX점'과 'XX' 두 가지로 섞여 나온다. 헷갈리지 말고 정식 매장명으로 통일한다.
- 스마트직영점은 근태를 안 올린다. 미오픈은 없어야 하고, 10시/11시 오픈 여부만 확인한다.
  → 스마트직영점이 안 올라왔다고 '누락'으로 보고하지 않는다.
- 휴무는 월 6회까지 가능, 조기마감/조기퇴근은 각 월 1회까지.
  이걸 넘기는 매장이 있으면 issues에 kind='규정위반'으로 알린다.
"""

RULES_BY_CATEGORY = {
    "본사": RULES_BONSA,
    "직영": RULES_JIKYEONG,
    "소사장": RULES_SOSAJANG,
}


# ---------------------------------------------------------------------------
# 구조화된 응답 스키마 (tool use로 강제해서 JSON 파싱 실패를 없앰)
# ---------------------------------------------------------------------------

REPORT_TOOL = {
    "name": "근태검토결과",
    "description": "카톡 대화를 검토한 결과를 구조화해서 보고합니다.",
    "input_schema": {
        "type": "object",
        "properties": {
            "records": {
                "type": "array",
                "description": "카톡에서 근거가 확인된 근태 기록만 넣습니다. 추측 금지.",
                "items": {
                    "type": "object",
                    "properties": {
                        "date": {"type": "string", "description": "YYYY-MM-DD"},
                        "target": {"type": "string", "description": "본사·직영은 직원 이름, 소사장은 매장명"},
                        "code": {"type": "string", "description": "근태 코드 (소사장은 출근보고 코드)"},
                        "close_code": {"type": "string", "description": "소사장 전용: 퇴근보고 코드. 해당 없으면 빈 문자열"},
                        "memo": {"type": "string", "description": "예외 문구나 사유 원문. 없으면 빈 문자열"},
                        "source": {"type": "string", "description": "어떤 메시지를 근거로 했는지 짧게"},
                        "confidence": {"type": "string", "enum": ["high", "medium", "low"]},
                    },
                    "required": ["date", "target", "code", "confidence"],
                },
            },
            "issues": {
                "type": "array",
                "description": "사람이 봐야 하는 것들. 오타/누락/중복/확인필요/규정위반.",
                "items": {
                    "type": "object",
                    "properties": {
                        "kind": {
                            "type": "string",
                            "enum": ["오타", "누락", "중복", "확인필요", "규정위반"],
                        },
                        "date": {"type": "string", "description": "관련 날짜. 없으면 빈 문자열"},
                        "target": {"type": "string", "description": "관련 직원명/매장명. 없으면 빈 문자열"},
                        "detail": {"type": "string", "description": "무엇이 문제인지 한국어로 구체적으로"},
                        "suggestion": {"type": "string", "description": "어떻게 하면 될지 제안. 확인필요면 질문 문장"},
                    },
                    "required": ["kind", "detail"],
                },
            },
            "summary": {"type": "string", "description": "전체 검토 요약 2~4문장, 한국어"},
        },
        "required": ["records", "issues", "summary"],
    },
}


def _build_system_prompt(category, roster, year, month):
    roster_text = "\n".join(f"  - {r}" for r in roster) if roster else "  (등록된 명단 없음)"
    label = "매장 목록" if category == "소사장" else "직원 명단"
    return f"""너는 (주)대교통신의 근태 담당자를 돕는 검토 보조자다.
카카오톡 단톡방 대화 내용을 읽고 {year}년 {month}월 근태를 정리하되, 무엇보다 **틀린 것을 찾아내는 게 네 역할**이다.

{RULES_COMMON}
{RULES_BY_CATEGORY.get(category, "")}

[이번에 검토할 {label} — 여기 없는 이름/매장이 나오면 오타이거나 새 인원이다]
{roster_text}

[반드시 지킬 것]
- 대상 기간은 {year}년 {month}월이다. 다른 달 내용은 무시한다.
- date는 반드시 YYYY-MM-DD 형식으로 쓴다.
- 확신이 없으면 records에 억지로 넣지 말고 issues에 넣어서 물어본다.
  담당자가 직접 확인하고 고칠 것이므로, 애매한 걸 조용히 넘기는 게 제일 나쁘다.
- 모든 설명(detail, suggestion, summary)은 한국어로, 실무자가 바로 알아들을 수 있게 쓴다.
- 반드시 '근태검토결과' 도구를 호출해서 답한다."""


def _split_chunks(text, max_chars=MAX_CHARS_PER_CHUNK):
    """카톡 대화를 날짜 구분선 기준으로 잘라서 여러 덩어리로 나눔"""
    if len(text) <= max_chars:
        return [text]
    lines = text.splitlines(keepends=True)
    # 카톡 내보내기의 날짜 구분선 (예: --------------- 2026년 7월 1일 수요일 ---------------)
    date_line = re.compile(r"^-{3,}.*\d{4}년\s*\d{1,2}월\s*\d{1,2}일.*-{3,}\s*$")
    chunks, cur, cur_len = [], [], 0
    for ln in lines:
        if cur_len + len(ln) > max_chars and cur and date_line.match(ln):
            chunks.append("".join(cur))
            cur, cur_len = [], 0
        elif cur_len + len(ln) > max_chars * 1.3 and cur:
            # 날짜 구분선을 못 찾을 만큼 길면 그냥 자름
            chunks.append("".join(cur))
            cur, cur_len = [], 0
        cur.append(ln)
        cur_len += len(ln)
    if cur:
        chunks.append("".join(cur))
    return chunks


def review_kakao(category, chat_text, roster, year, month, model=None, progress=None):
    """카톡 대화를 검토해서 {'records': [...], 'issues': [...], 'summary': str} 반환"""
    model = model or DEFAULT_MODEL
    client = _client()
    system = _build_system_prompt(category, roster, year, month)
    chunks = _split_chunks(chat_text)

    all_records, all_issues, summaries = [], [], []
    for i, chunk in enumerate(chunks, 1):
        if progress:
            progress(i, len(chunks))
        part_note = ""
        if len(chunks) > 1:
            part_note = (
                f"\n\n(참고: 대화가 길어서 {len(chunks)}개로 나눠 보내고 있다. "
                f"지금은 {i}번째 조각이다. 이 조각에서 확인되는 것만 보고해라.)"
            )
        resp = client.messages.create(
            model=model,
            max_tokens=8000,
            system=system,
            tools=[REPORT_TOOL],
            tool_choice={"type": "tool", "name": "근태검토결과"},
            messages=[{
                "role": "user",
                "content": f"아래는 카카오톡 대화 내용이다. 검토해줘.{part_note}\n\n<대화내용>\n{chunk}\n</대화내용>",
            }],
        )
        payload = None
        for block in resp.content:
            if getattr(block, "type", None) == "tool_use":
                payload = block.input
                break
        if not payload:
            continue
        all_records.extend(payload.get("records") or [])
        all_issues.extend(payload.get("issues") or [])
        if payload.get("summary"):
            summaries.append(payload["summary"])

    # 조각을 나눠 보냈을 때 생기는 중복 제거
    seen, records = set(), []
    for r in all_records:
        key = (r.get("date"), r.get("target"), r.get("code"), r.get("close_code"))
        if key in seen:
            continue
        seen.add(key)
        records.append(r)
    records.sort(key=lambda r: (r.get("date") or "", r.get("target") or ""))

    seen_i, issues = set(), []
    for it in all_issues:
        key = (it.get("kind"), it.get("date"), it.get("target"), it.get("detail"))
        if key in seen_i:
            continue
        seen_i.add(key)
        issues.append(it)

    return {
        "category": category,
        "year": year,
        "month": month,
        "records": records,
        "issues": issues,
        "summary": " ".join(summaries) if summaries else "",
    }


def to_json(result):
    return json.dumps(result, ensure_ascii=False)


def from_json(s):
    return json.loads(s)
