# -*- coding: utf-8 -*-
"""
ai_engine.py
OpenAI API를 호출하여
  1) 지문에서 Vocabulary List 후보를 생성하고
  2) 승인된 Vocabulary List를 기준으로 MASTER MANUAL v1.0 5장 규격의
     30문항 Vocabulary Review Test를 생성한다.

원칙 (MASTER MANUAL 1장): AI는 내용 데이터를 생성하고, 문서 디자인은
Word 템플릿/Python 코드(docx_builder.py)가 담당한다. 이 모듈은 절대
문서 레이아웃을 다루지 않고 구조화된 JSON 데이터만 만든다.
"""

import json
import re

from openai import OpenAI, APIConnectionError, APIStatusError


class AIEngineError(Exception):
    """AI 호출/응답 파싱 중 발생한 오류."""

    def __init__(self, message: str, raw_response: str = ""):
        super().__init__(message)
        self.raw_response = raw_response


DIFFICULTY_GUIDE = {
    "Easier": (
        "쉬움: 문맥 단서를 더 제공하고 객관식 오답의 혼동도를 낮춘다. "
        "Reading Review 문항은 원문 상황과 더 가깝게 만든다."
    ),
    "Normal": (
        "보통: 문맥 단서, 오답 매력도, 문장 난이도를 균형 있게 유지한다."
    ),
    "Harder": (
        "어려움: 문맥 단서를 일부 줄이고 오답을 더 그럴듯하게 만들되, "
        "원 지문 수준을 크게 넘어서는 문법으로 난도를 올리지 않는다."
    ),
}


def get_client(api_key: str) -> OpenAI:
    if not api_key or not api_key.strip():
        raise AIEngineError("OpenAI API 키가 설정되지 않았습니다.")
    return OpenAI(api_key=api_key.strip())


def _extract_json(text: str):
    """모델 응답 텍스트에서 JSON 부분만 강건하게 추출한다."""
    text = text.strip()

    fence_match = re.search(r"```(?:json)?\s*(.*?)\s*```", text, re.DOTALL)
    candidate = fence_match.group(1).strip() if fence_match else text

    try:
        return json.loads(candidate)
    except json.JSONDecodeError:
        pass

    start_idx = None
    for i, ch in enumerate(candidate):
        if ch in "[{":
            start_idx = i
            break
    if start_idx is None:
        raise AIEngineError("AI 응답에서 JSON 데이터를 찾을 수 없습니다.", raw_response=text)

    open_ch = candidate[start_idx]
    close_ch = "]" if open_ch == "[" else "}"
    depth = 0
    for j in range(start_idx, len(candidate)):
        if candidate[j] == open_ch:
            depth += 1
        elif candidate[j] == close_ch:
            depth -= 1
            if depth == 0:
                snippet = candidate[start_idx:j + 1]
                try:
                    return json.loads(snippet)
                except json.JSONDecodeError as e:
                    raise AIEngineError(f"AI 응답 JSON 파싱에 실패했습니다: {e}", raw_response=text)

    raise AIEngineError("AI 응답에서 완전한 JSON을 찾지 못했습니다 (응답이 잘렸을 수 있습니다).", raw_response=text)


def _call_openai(client, model, system_prompt, user_prompt, max_tokens):
    """OpenAI Responses API를 호출하고 텍스트 응답을 반환한다."""
    try:
        response = client.responses.create(
            model=model,
            instructions=system_prompt,
            input=user_prompt,
            max_output_tokens=max_tokens,
        )
    except APIStatusError as e:
        status = getattr(e, "status_code", "unknown")
        message = getattr(e, "message", str(e))
        raise AIEngineError(f"OpenAI API 오류 (status {status}): {message}")
    except APIConnectionError as e:
        raise AIEngineError(f"OpenAI API 연결 오류: {e}")
    except Exception as e:
        raise AIEngineError(f"OpenAI API 호출 중 오류: {e}")

    full_text = (getattr(response, "output_text", "") or "").strip()
    if not full_text:
        raise AIEngineError("AI가 빈 응답을 반환했습니다.")
    return full_text


VOCAB_SYSTEM_PROMPT = """당신은 한국 중고등학생 대상 영어 독해 수업을 오래 담당해 온 전문 영어 교사이자
어휘 교재 편집자입니다. 주어진 영어 지문을 분석하여 Vocabulary List를 만듭니다.

반드시 지켜야 할 선정 원칙:
1. 단어 개수를 먼저 정해 놓고 억지로 채우지 않는다. 지문 분량과 난이도에 맞는 만큼만 선정한다.
2. 지문 이해에 중요한 어휘, 그리고 다른 독해 지문에서도 재사용 가치가 높은 어휘를 우선한다.
3. 단일 단어뿐 아니라 phrasal verb, idiom, collocation, expression도 학습 가치가 있으면 하나의 항목으로 선정한다.
4. 고유명사 및 학습 가치가 낮은 지나치게 쉬운 단어는 원칙적으로 제외한다.
5. 지문에 실제로 등장하거나 지문에서 명확히 확인 가능한 표현만 근거로 삼는다. 지문에 없는 단어를 지어내지 않는다.

각 어휘 항목은 다음 필드를 모두 포함해야 합니다:
- word: 학습 대상 단어 또는 표현 (지문에 등장한 원형 그대로)
- pronunciation: 발음 정보 (가능하면 IPA 표기, 일관된 형식으로)
- pos: 품사 (n., v., adj., adv., phrasal v., idiom 등 약어로)
- meaning: 한국어 의미. 사전 전체 뜻이 아니라 해당 지문 문맥에서의 의미 중심으로 작성
- synonym: 동의어. 정확한 의미 관계가 있을 때만 작성하고, 없으면 반드시 "-"
- antonym: 반의어. 정확한 의미 관계가 있을 때만 작성하고, 없으면 반드시 "-"
- new_example: 이 단어를 사용한 새 예문(영어). 지문 원문을 복사하지 않고 새로 작성. 이후 시험 문제에 그대로 재사용되지 않을 문장이어야 함
- source: 지문 내 근거 위치 (예: "1문단", "Page 2 3번째 문장" 등 추적 가능하게)
- importance: 1~5 사이의 정수. 5가 가장 중요.

반드시 JSON 배열만 출력하세요. 다른 설명, 인사말, 코드블록 표시(```) 없이 순수 JSON 배열 텍스트만 출력합니다.
예시 형식:
[{"word":"...", "pronunciation":"...", "pos":"...", "meaning":"...", "synonym":"...", "antonym":"...", "new_example":"...", "source":"...", "importance":3}]
"""


def extract_vocabulary(client, model, title, passage_text, target_count_hint=0):
    hint_line = ""
    if target_count_hint and target_count_hint > 0:
        hint_line = (
            f"\n참고용 목표 단어 수는 {target_count_hint}개이지만, 이는 강제 기준이 아닙니다. "
            "학습 가치가 있는 만큼만 선정하고 억지로 개수를 맞추지 마세요."
        )

    user_prompt = f"""자료 제목: {title}
{hint_line}

아래는 분석 대상 지문입니다 ([Page N] 표시는 원문 페이지 번호입니다):
---
{passage_text}
---

위 지문을 분석하여 Vocabulary List를 JSON 배열로만 출력하세요."""

    raw = _call_openai(client, model, VOCAB_SYSTEM_PROMPT, user_prompt, max_tokens=6000)
    data = _extract_json(raw)

    if not isinstance(data, list):
        raise AIEngineError("Vocabulary List 응답이 JSON 배열 형식이 아닙니다.", raw_response=raw)

    normalized = []
    for item in data:
        if not isinstance(item, dict):
            continue
        normalized.append({
            "word": str(item.get("word", "")).strip(),
            "pronunciation": str(item.get("pronunciation", "")).strip(),
            "pos": str(item.get("pos", "")).strip(),
            "meaning": str(item.get("meaning", "")).strip(),
            "synonym": (str(item.get("synonym", "")).strip() or "-"),
            "antonym": (str(item.get("antonym", "")).strip() or "-"),
            "new_example": str(item.get("new_example", "")).strip(),
            "source": str(item.get("source", "")).strip(),
            "importance": item.get("importance", 3),
        })

    if not normalized:
        raise AIEngineError("AI가 유효한 어휘 항목을 반환하지 않았습니다.", raw_response=raw)

    return normalized


TEST_SYSTEM_PROMPT = """당신은 Reading Vocabulary Maker 시스템의 시험 출제 엔진입니다.
승인된 Vocabulary List와 원 지문을 바탕으로 정확히 "30문항"의 Vocabulary Review Test를
아래 MASTER MANUAL v1.0 규격에 맞춰 생성합니다.

[문항 구성 - 반드시 정확히 지킬 것]
- 1~12번 (12문항): type="reading_review"
  - 원 지문의 핵심 사실/상황/원인·결과/대조 관계를 보존한 새로운 문장을 만들고, 빈칸 + 첫 글자 힌트 형식으로 출제.
  - 원문 문장을 그대로 복사하고 단어 하나만 지우는 방식 절대 금지.
  - 학생이 원 지문을 공부했다면 알아볼 수 있는 내용이어야 함.
  - question_text 예시 형식: "The scientists were _______ (a______) by the unexpected results because ..."
    빈칸은 "_______" 로, 첫 글자 힌트는 그 뒤 괄호 안에 "(첫글자______)" 형식으로 넣는다.
- 13~21번 (9문항): type="vocabulary_transfer"
  - 원 지문과 다른 새로운 문맥의 문장을 새로 작성. 원문 문장 그대로 사용 금지.
  - Vocabulary List의 new_example 문장을 그대로 재사용 금지 (반드시 또 다른 새 문장).
  - 마찬가지로 빈칸 + 첫 글자 힌트 형식.
- 22~26번 (5문항): type="english_definition"
  - 영어 정의를 읽고 목표 어휘를 고르는 4지선다형. question_text에 영어 정의를 넣는다.
  - choices는 4개, 그 중 정답 하나만 정확히 성립해야 하며 오답은 품사/난이도가 비슷하되 의미는 분명히 달라야 함.
  - 정의가 목표 단어를 그대로 노출하면 안 됨.
- 27~30번 (4문항): type="vocabulary_relations"
  - 동의어/반의어 등 정확한 어휘 관계를 확인하는 4지선다형.
  - question_text에 "다음 단어의 동의어(또는 반의어)로 가장 적절한 것은?  target: <word>" 형식처럼 무엇을 묻는지 명확히 표시.
  - choices 4개 중 정답은 정확한 관계가 성립하는 것 하나뿐이어야 하며, 부정확한 synonym/antonym을 만들어내지 않는다.

[공통 규칙]
- 1~21번의 target_word는 반드시 승인된 Vocabulary List에 있는 word 값과 정확히 일치해야 한다 (대소문자 무관).
- 같은 목표 어휘를 시험 전체에서 불필요하게 반복하지 않는다 (부득이한 경우가 아니면 1회씩만 사용).
- Vocabulary List의 순서를 그대로 시험 문항 순서로 기계적으로 사용하지 않는다.
- 지문에 없는 사실을 만들어 문제에 넣지 않는다.
- 모든 객관식(22~30번)은 정답이 단 하나만 존재해야 한다.
- 각 문항은 answer(정답 텍스트), explanation(정답 근거를 학생이 이해할 수 있도록 짧고 정확하게 설명, 2~3문장 이내)을 포함해야 한다.
- English Definition/Vocabulary Relations 문항은 choices(4개 문자열 배열)와 answer_index(0~3, 정답의 인덱스)를 포함해야 한다.
- Reading Review/Vocabulary Transfer 문항은 choices, answer_index를 생략(null)하고 answer에 정답 단어/표현 텍스트를 넣는다.

[난이도 지침]
{difficulty_guide}

각 문항은 다음 필드를 가진 객체입니다:
{{"no": 정수, "type": "reading_review|vocabulary_transfer|english_definition|vocabulary_relations",
  "question_text": "...", "choices": ["...","...","...","..."] 또는 null,
  "answer": "...", "answer_index": 정수 또는 null, "target_word": "...", "explanation": "..."}}

반드시 다음 형식의 JSON 객체만 출력하세요. 다른 설명, 인사말, 코드블록 표시(```) 없이 순수 JSON만 출력합니다:
{{"questions": [ ... 정확히 30개 ... ]}}
"""


def generate_test(client, model, title, passage_text, approved_vocab_list, difficulty, retry_issues=None):
    difficulty_guide = DIFFICULTY_GUIDE.get(difficulty, DIFFICULTY_GUIDE["Normal"])
    system_prompt = TEST_SYSTEM_PROMPT.format(difficulty_guide=difficulty_guide)

    vocab_json = json.dumps(
        [
            {
                "word": v["word"],
                "pos": v["pos"],
                "meaning": v["meaning"],
                "synonym": v.get("synonym", "-"),
                "antonym": v.get("antonym", "-"),
                "new_example": v.get("new_example", ""),
            }
            for v in approved_vocab_list
        ],
        ensure_ascii=False,
    )

    retry_note = ""
    if retry_issues:
        issues_text = "\n".join(f"- {issue}" for issue in retry_issues)
        retry_note = f"""

[이전 생성 결과의 자동 검수 오류 - 반드시 모두 수정할 것]
{issues_text}
위 문제들을 모두 해결한 새로운 30문항 세트를 처음부터 다시 생성하세요."""

    user_prompt = f"""자료 제목: {title}

[승인된 Vocabulary List (JSON)]
{vocab_json}

[원 지문]
---
{passage_text}
---
{retry_note}

위 조건에 맞춰 정확히 30문항의 JSON을 생성하세요."""

    raw = _call_openai(client, model, system_prompt, user_prompt, max_tokens=8000)
    data = _extract_json(raw)

    if not isinstance(data, dict) or "questions" not in data:
        raise AIEngineError("시험 응답이 예상된 JSON 형식({'questions': [...]})이 아닙니다.", raw_response=raw)

    questions = data["questions"]
    if not isinstance(questions, list):
        raise AIEngineError("questions 필드가 배열이 아닙니다.", raw_response=raw)

    normalized = []
    for q in questions:
        if not isinstance(q, dict):
            continue
        try:
            no = int(q.get("no"))
        except (TypeError, ValueError):
            no = None
        choices = q.get("choices")
        if isinstance(choices, list):
            choices = [str(c).strip() for c in choices]
        else:
            choices = None
        answer_index = q.get("answer_index")
        try:
            answer_index = int(answer_index) if answer_index is not None else None
        except (TypeError, ValueError):
            answer_index = None

        normalized.append({
            "no": no,
            "type": str(q.get("type", "")).strip(),
            "question_text": str(q.get("question_text", "")).strip(),
            "choices": choices,
            "answer": str(q.get("answer", "")).strip(),
            "answer_index": answer_index,
            "target_word": str(q.get("target_word", "")).strip(),
            "explanation": str(q.get("explanation", "")).strip(),
        })

    return normalized
