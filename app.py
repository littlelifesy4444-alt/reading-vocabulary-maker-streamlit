# -*- coding: utf-8 -*-
"""
Reading Vocabulary Maker
=========================
MASTER MANUAL v1.0 기준 전체 제작 흐름을 구현한 Streamlit 앱.

PDF 업로드 → PDF 분석 → Vocabulary List 생성(AI) → 사용자 검토/수정/삭제/추가
→ 승인 → 30문항 시험 생성(AI) → 자동 검수 → Word 3종 생성 → ZIP 다운로드

실행 방법:
    streamlit run app.py
"""

import io
import zipfile
from datetime import datetime

import pandas as pd
import streamlit as st

import ai_engine
import docx_builder
import pdf_utils
import validators


# ---------------------------------------------------------------------------
# 페이지 설정 & 세션 상태 초기화
# ---------------------------------------------------------------------------

st.set_page_config(page_title="Reading Vocabulary Maker", page_icon="📚", layout="centered")

MODEL_OPTIONS = ["gpt-5.6", "gpt-5.6-terra", "gpt-5.6-luna"]

DEFAULTS = {
    "api_key": "",
    "model": MODEL_OPTIONS[0],
    "title": "",
    "difficulty": "Normal",
    "target_count_hint": 0,
    "pdf_bytes": None,
    "pdf_name": None,
    "pdf_total_pages": 0,
    "start_page": 1,
    "end_page": 1,
    "passage_text": "",
    "extraction_done": False,
    "vocab_list": [],
    "vocab_approved": False,
    "vocab_issues": [],
    "test_questions": [],
    "test_issues": [],
    "test_generated": False,
    "last_raw_response": "",
    "final_zip": None,
    "final_zip_name": None,
}

for k, v in DEFAULTS.items():
    if k not in st.session_state:
        st.session_state[k] = v


def reset_all():
    keep_keys = {"api_key", "model"}
    for k, v in DEFAULTS.items():
        if k not in keep_keys:
            st.session_state[k] = v


VOCAB_COLUMNS = ["word", "pronunciation", "pos", "meaning", "synonym", "antonym",
                  "new_example", "source", "importance"]
VOCAB_COLUMN_LABELS = {
    "word": "Word / Expression",
    "pronunciation": "Pronunciation",
    "pos": "POS",
    "meaning": "Meaning (한국어)",
    "synonym": "Synonym",
    "antonym": "Antonym",
    "new_example": "New Example",
    "source": "Source",
    "importance": "Importance (1-5)",
}


def vocab_list_to_df(vocab_list):
    if not vocab_list:
        return pd.DataFrame([{c: "" for c in VOCAB_COLUMNS}]).iloc[0:0]
    rows = []
    for item in vocab_list:
        row = {c: item.get(c, "") for c in VOCAB_COLUMNS}
        try:
            row["importance"] = int(row["importance"])
        except (TypeError, ValueError):
            row["importance"] = 3
        rows.append(row)
    return pd.DataFrame(rows, columns=VOCAB_COLUMNS)


def df_to_vocab_list(df):
    vocab_list = []
    for _, row in df.iterrows():
        word = str(row.get("word", "")).strip()
        if not word:
            continue
        importance = row.get("importance", 3)
        try:
            importance = int(importance)
        except (TypeError, ValueError):
            importance = 3
        vocab_list.append({
            "word": word,
            "pronunciation": str(row.get("pronunciation", "")).strip(),
            "pos": str(row.get("pos", "")).strip(),
            "meaning": str(row.get("meaning", "")).strip(),
            "synonym": str(row.get("synonym", "-")).strip() or "-",
            "antonym": str(row.get("antonym", "-")).strip() or "-",
            "new_example": str(row.get("new_example", "")).strip(),
            "source": str(row.get("source", "")).strip(),
            "importance": importance,
        })
    return vocab_list


# ---------------------------------------------------------------------------
# 사이드바 - API 설정
# ---------------------------------------------------------------------------

# OpenAI API key: Streamlit Secrets의 OPENAI_API_KEY 사용
try:
    st.session_state["api_key"] = st.secrets.get("OPENAI_API_KEY", "")
except Exception:
    st.session_state["api_key"] = ""

with st.sidebar:
    st.header("⚙️ AI 설정")
    if st.session_state["api_key"]:
        st.success("OpenAI API 키 연결됨")
    else:
        st.error("Streamlit Secrets에 OPENAI_API_KEY를 설정해주세요.")
    st.session_state["model"] = st.selectbox(
        "모델", MODEL_OPTIONS,
        index=MODEL_OPTIONS.index(st.session_state["model"])
        if st.session_state["model"] in MODEL_OPTIONS else 0,
    )
    st.divider()
    if st.button("🔄 처음부터 다시 시작", use_container_width=True):
        reset_all()
        st.rerun()

st.title("📚 Reading Vocabulary Maker")
st.caption("MASTER MANUAL v1.0 기준: 독해 → 어휘 선정 → Vocabulary List → 시험 → 정답·해설")


# ---------------------------------------------------------------------------
# STEP 1. 기본 정보
# ---------------------------------------------------------------------------

st.header("1. 기본 정보 입력")
st.session_state["title"] = st.text_input(
    "교재명 / Chapter명 / 시험지 제목",
    value=st.session_state["title"],
    placeholder="예) High School Reading Ch.5 - The Ocean's Hidden World",
)
st.session_state["difficulty"] = st.radio(
    "난이도", ["Easier", "Normal", "Harder"],
    index=["Easier", "Normal", "Harder"].index(st.session_state["difficulty"]),
    horizontal=True,
)
st.session_state["target_count_hint"] = st.number_input(
    "참고용 목표 어휘 수 (0 = AI가 자동 판단, 강제 아님)",
    min_value=0, max_value=60, value=int(st.session_state["target_count_hint"]), step=1,
)


# ---------------------------------------------------------------------------
# STEP 2. PDF 업로드 & 페이지 범위
# ---------------------------------------------------------------------------

st.header("2. 독해 PDF 업로드")
uploaded_file = st.file_uploader("PDF 파일 선택", type=["pdf"], accept_multiple_files=False)

if uploaded_file is not None:
    file_bytes = uploaded_file.getvalue()
    if st.session_state["pdf_bytes"] != file_bytes:
        st.session_state["pdf_bytes"] = file_bytes
        st.session_state["pdf_name"] = uploaded_file.name
        try:
            st.session_state["pdf_total_pages"] = pdf_utils.get_pdf_page_count(file_bytes)
            st.session_state["start_page"] = 1
            st.session_state["end_page"] = st.session_state["pdf_total_pages"]
        except pdf_utils.PdfExtractionError as e:
            st.error(str(e))
            st.session_state["pdf_total_pages"] = 0

if st.session_state["pdf_bytes"] and st.session_state["pdf_total_pages"] > 0:
    st.success(f"업로드됨: {st.session_state['pdf_name']} (총 {st.session_state['pdf_total_pages']}페이지)")
    col1, col2 = st.columns(2)
    with col1:
        st.session_state["start_page"] = st.number_input(
            "시작 페이지", min_value=1, max_value=st.session_state["pdf_total_pages"],
            value=min(st.session_state["start_page"], st.session_state["pdf_total_pages"]),
        )
    with col2:
        st.session_state["end_page"] = st.number_input(
            "끝 페이지", min_value=1, max_value=st.session_state["pdf_total_pages"],
            value=min(st.session_state["end_page"], st.session_state["pdf_total_pages"]),
        )


# ---------------------------------------------------------------------------
# STEP 3. 분석 실행 (Vocabulary List 생성)
# ---------------------------------------------------------------------------

st.header("3. 어휘 분석 실행")

analyze_disabled = not (
    st.session_state["api_key"]
    and st.session_state["title"].strip()
    and st.session_state["pdf_bytes"]
    and st.session_state["pdf_total_pages"] > 0
)
if analyze_disabled:
    st.info("OpenAI API 키, 제목, PDF 업로드를 먼저 완료해주세요.")

if st.button("🔍 PDF 분석 및 Vocabulary List 생성", disabled=analyze_disabled, use_container_width=True):
    with st.spinner("PDF에서 텍스트를 추출하는 중..."):
        try:
            text, total_pages, pages_with_text = pdf_utils.extract_text(
                st.session_state["pdf_bytes"],
                start_page=st.session_state["start_page"],
                end_page=st.session_state["end_page"],
            )
            ai_text, truncated = pdf_utils.truncate_text_for_ai(text)
            st.session_state["passage_text"] = ai_text
            if truncated:
                st.warning("지문이 매우 길어 AI 분석용 텍스트를 일부 구간까지만 사용했습니다 (문서 생성에는 영향 없음).")
        except pdf_utils.PdfExtractionError as e:
            st.error(str(e))
            st.stop()

    with st.spinner("AI가 Vocabulary List를 생성하는 중... (지문 길이에 따라 다소 시간이 걸릴 수 있습니다)"):
        try:
            client = ai_engine.get_client(st.session_state["api_key"])
            vocab_list = ai_engine.extract_vocabulary(
                client,
                st.session_state["model"],
                st.session_state["title"],
                st.session_state["passage_text"],
                target_count_hint=st.session_state["target_count_hint"],
            )
            st.session_state["vocab_list"] = vocab_list
            st.session_state["extraction_done"] = True
            st.session_state["vocab_approved"] = False
            st.session_state["test_generated"] = False
            st.session_state["test_questions"] = []
            st.success(f"Vocabulary List {len(vocab_list)}개 항목이 생성되었습니다. 아래에서 검토해주세요.")
        except ai_engine.AIEngineError as e:
            st.error(f"Vocabulary List 생성 실패: {e}")
            if e.raw_response:
                with st.expander("AI 원본 응답 (디버그용)"):
                    st.code(e.raw_response)


# ---------------------------------------------------------------------------
# STEP 4. Vocabulary List 검토 / 수정 / 삭제 / 추가
# ---------------------------------------------------------------------------

if st.session_state["extraction_done"]:
    st.header("4. Vocabulary List 검토 · 수정 · 삭제 · 추가")
    st.caption("표를 직접 수정하세요. 행을 삭제하려면 왼쪽 체크박스를 선택 후 delete 키, 새 행을 추가하려면 표 맨 아래 빈 행을 사용하세요.")

    df = vocab_list_to_df(st.session_state["vocab_list"])
    edited_df = st.data_editor(
        df,
        num_rows="dynamic",
        use_container_width=True,
        key="vocab_editor",
        column_config={
            "word": st.column_config.TextColumn(VOCAB_COLUMN_LABELS["word"], required=True),
            "pronunciation": st.column_config.TextColumn(VOCAB_COLUMN_LABELS["pronunciation"]),
            "pos": st.column_config.TextColumn(VOCAB_COLUMN_LABELS["pos"]),
            "meaning": st.column_config.TextColumn(VOCAB_COLUMN_LABELS["meaning"]),
            "synonym": st.column_config.TextColumn(VOCAB_COLUMN_LABELS["synonym"]),
            "antonym": st.column_config.TextColumn(VOCAB_COLUMN_LABELS["antonym"]),
            "new_example": st.column_config.TextColumn(VOCAB_COLUMN_LABELS["new_example"], width="large"),
            "source": st.column_config.TextColumn(VOCAB_COLUMN_LABELS["source"]),
            "importance": st.column_config.NumberColumn(VOCAB_COLUMN_LABELS["importance"], min_value=1, max_value=5, step=1),
        },
    )

    col_a, col_b = st.columns(2)
    with col_a:
        run_check = st.button("✅ 검수 실행", use_container_width=True)
    with col_b:
        approve = st.button("🔒 최종 어휘 승인", use_container_width=True, type="primary")

    current_vocab_list = df_to_vocab_list(edited_df)

    if run_check or approve:
        st.session_state["vocab_issues"] = validators.validate_vocab_list(current_vocab_list)
        st.session_state["vocab_list"] = current_vocab_list

    if st.session_state["vocab_issues"]:
        st.error("검수 결과: 다음 문제를 해결해야 합니다.")
        for issue in st.session_state["vocab_issues"]:
            st.markdown(f"- {issue}")
    elif run_check:
        st.success("검수 통과: 문제가 발견되지 않았습니다.")

    if approve:
        if st.session_state["vocab_issues"]:
            st.error("검수 문제를 먼저 해결한 뒤 승인할 수 있습니다.")
        elif not current_vocab_list:
            st.error("승인할 어휘가 없습니다.")
        else:
            st.session_state["vocab_list"] = current_vocab_list
            st.session_state["vocab_approved"] = True
            st.session_state["test_generated"] = False
            st.session_state["test_questions"] = []
            st.success(f"Vocabulary List {len(current_vocab_list)}개 항목이 승인되었습니다.")

    if st.session_state["vocab_approved"]:
        st.info(f"✔ 현재 승인된 Vocabulary List: {len(st.session_state['vocab_list'])}개 항목. "
                "다시 수정하려면 표를 편집한 뒤 '최종 어휘 승인'을 다시 눌러주세요.")


# ---------------------------------------------------------------------------
# STEP 5. 시험 생성 (승인된 어휘 기준)
# ---------------------------------------------------------------------------

if st.session_state["vocab_approved"]:
    st.header("5. Vocabulary Review Test 생성 (30문항)")
    st.caption("Reading Review 1-12 / Vocabulary Transfer 13-21 / English Definition 22-26 / Vocabulary Relations 27-30")

    gen_col1, gen_col2 = st.columns(2)
    with gen_col1:
        gen_test_clicked = st.button("📝 30문항 시험 생성", use_container_width=True)
    with gen_col2:
        regen_test_clicked = st.button(
            "♻️ 검수 오류 반영하여 재생성",
            use_container_width=True,
            disabled=not st.session_state["test_issues"],
        )

    if gen_test_clicked or regen_test_clicked:
        retry_issues = st.session_state["test_issues"] if regen_test_clicked else None
        with st.spinner("AI가 30문항 시험을 생성하는 중... (다소 시간이 걸릴 수 있습니다)"):
            try:
                client = ai_engine.get_client(st.session_state["api_key"])
                questions = ai_engine.generate_test(
                    client,
                    st.session_state["model"],
                    st.session_state["title"],
                    st.session_state["passage_text"],
                    st.session_state["vocab_list"],
                    st.session_state["difficulty"],
                    retry_issues=retry_issues,
                )
                st.session_state["test_questions"] = questions
                st.session_state["test_generated"] = True
                issues = validators.validate_test(questions, st.session_state["vocab_list"])
                st.session_state["test_issues"] = issues
                if issues:
                    st.warning(f"시험이 생성되었지만 자동 검수에서 {len(issues)}건의 문제가 발견되었습니다. 아래를 확인하세요.")
                else:
                    st.success("시험 생성 및 자동 검수 통과!")
            except ai_engine.AIEngineError as e:
                st.error(f"시험 생성 실패: {e}")
                if e.raw_response:
                    with st.expander("AI 원본 응답 (디버그용)"):
                        st.code(e.raw_response)


# ---------------------------------------------------------------------------
# STEP 6. 자동 검수 결과 & 미리보기
# ---------------------------------------------------------------------------

if st.session_state["test_generated"] and st.session_state["test_questions"]:
    st.header("6. 자동 검수 체크리스트")

    checklist_labels = [
        "중복 어휘 없음", "필수 필드 존재", "발음·품사 존재", "정확히 30문항",
        "번호·유형 배치(1-12/13-21/22-26/27-30) 정확", "1-21번 target word가 승인 목록에 존재",
        "22-30번 선택지 4개", "객관식 정답 단일성", "New Example 재사용 없음", "시험지-정답지 1:1 대응",
    ]
    if not st.session_state["test_issues"] and not st.session_state["vocab_issues"]:
        st.success("✅ 모든 자동 검수 항목을 통과했습니다. 최종 문서를 생성할 수 있습니다.")
    else:
        st.error(f"❌ 검수 미통과 항목이 있습니다 ({len(st.session_state['test_issues'])}건). "
                 "최종 문서를 생성하기 전에 '검수 오류 반영하여 재생성'을 실행하거나 어휘 목록을 조정해주세요.")
        for issue in st.session_state["test_issues"]:
            st.markdown(f"- {issue}")

    with st.expander("📋 생성된 시험 문항 미리보기"):
        for q in sorted(st.session_state["test_questions"], key=lambda x: (x.get("no") is None, x.get("no"))):
            st.markdown(f"**{q.get('no')}. [{q.get('type')}]** {q.get('question_text')}")
            if q.get("choices"):
                for i, c in enumerate(q["choices"]):
                    marker = "✅" if q.get("answer_index") == i else "　"
                    st.markdown(f"&nbsp;&nbsp;{marker} {docx_builder.CHOICE_LABELS[i] if i < 4 else i} {c}")
            else:
                st.markdown(f"&nbsp;&nbsp;정답: `{q.get('answer')}`")
            st.caption(q.get("explanation", ""))


# ---------------------------------------------------------------------------
# STEP 7. 최종 Word 문서 생성 & ZIP 다운로드
# ---------------------------------------------------------------------------

if st.session_state["test_generated"] and st.session_state["test_questions"]:
    st.header("7. 최종 문서 생성 및 다운로드")

    ready_for_final = (not st.session_state["test_issues"]) and (not st.session_state["vocab_issues"])

    if not ready_for_final:
        st.warning("자동 검수를 통과해야 최종 문서를 생성할 수 있습니다 (MASTER MANUAL 8장/11장 원칙).")

    if st.button("📄 Word 문서 3종 생성 (Vocabulary List / Test / Answer&Explanation)",
                  disabled=not ready_for_final, use_container_width=True, type="primary"):
        with st.spinner("Word 문서를 생성하는 중..."):
            title = st.session_state["title"]
            vocab_list = st.session_state["vocab_list"]
            questions = st.session_state["test_questions"]
            difficulty = st.session_state["difficulty"]

            vocab_docx = docx_builder.build_vocab_list_docx(title, vocab_list)
            test_docx = docx_builder.build_test_docx(title, questions, difficulty)
            answer_docx = docx_builder.build_answer_docx(title, questions, vocab_list)

            safe_title = "".join(c for c in title if c.isalnum() or c in " _-").strip() or "Reading_Vocabulary_Maker"
            safe_title = safe_title.replace(" ", "_")

            zip_buffer = io.BytesIO()
            with zipfile.ZipFile(zip_buffer, "w", zipfile.ZIP_DEFLATED) as zf:
                zf.writestr(f"{safe_title}_Vocabulary_List.docx", vocab_docx.getvalue())
                zf.writestr(f"{safe_title}_Vocabulary_Test.docx", test_docx.getvalue())
                zf.writestr(f"{safe_title}_Answer_Explanation.docx", answer_docx.getvalue())
            zip_buffer.seek(0)

            st.session_state["final_zip"] = zip_buffer.getvalue()
            st.session_state["final_zip_name"] = f"{safe_title}_{datetime.now().strftime('%Y%m%d_%H%M')}.zip"
            st.success("Word 문서 3종 생성 완료! 아래 버튼으로 다운로드하세요.")

    if st.session_state.get("final_zip"):
        st.download_button(
            "⬇️ ZIP 다운로드",
            data=st.session_state["final_zip"],
            file_name=st.session_state["final_zip_name"],
            mime="application/zip",
            use_container_width=True,
        )
