from __future__ import annotations

import os
import shutil

import pandas as pd
import streamlit as st

import core


def secret(name: str, default: str = "") -> str:
    try:
        return st.secrets.get(name, os.getenv(name, default))
    except Exception:
        return os.getenv(name, default)


st.set_page_config(page_title="Reading Vocabulary Maker", page_icon="📘", layout="wide")
st.title("📘 Reading Vocabulary Maker")
st.caption("AI는 내용만 생성하고, 문서 디자인은 고정 Word 템플릿으로 잠그는 버전")

with st.expander("이 버전의 핵심", expanded=False):
    st.markdown(
        """
- **어휘 개수 고정 없음**: 실제 지문에서 필요한 만큼만 선별
- **2단계 생성**: 먼저 Vocabulary List를 검토·수정 → 승인 후 시험/해설 생성
- **고정 문서 템플릿**: 글꼴·칸 너비·세로 페이지·여백·footer를 AI가 건드리지 않음
- **자동 검수 통과 후에만 출력**: 30문항 구조, 정답 대응, New Example 재사용 등을 검사
"""
    )

with st.container(border=True):
    c1, c2 = st.columns([1.4, 1])
    with c1:
        title = st.text_input("교재/Chapter 제목", value=st.session_state.get("title", "Hackers TEPS Reading - Chapter 15"))
        unit_title = st.text_input("단원명", value=st.session_state.get("unit_title", "단원명"))
    with c2:
        page_range = st.text_input("PDF 내부 페이지 범위 (선택)", placeholder="예: 35-38 또는 35,37-38")
        model = st.text_input("OpenAI model", value=secret("OPENAI_MODEL", core.MODEL_DEFAULT))
    uploads = st.file_uploader(
        "독해 자료 PDF / 이미지",
        type=["pdf", "jpg", "jpeg", "png", "webp"],
        accept_multiple_files=True,
        help="PDF 1개를 올린 경우에만 페이지 범위를 지정할 수 있습니다.",
    )

if st.button("1️⃣ 어휘 분석", type="primary", use_container_width=True):
    if not uploads:
        st.warning("먼저 PDF 또는 이미지를 올려 주세요.")
    else:
        try:
            with st.spinner("원문을 읽고 학습 가치가 있는 어휘만 선별하고 있습니다..."):
                prepared = core.prepare_sources(uploads, page_range)
                result = core.analyze_vocabulary(prepared, title.strip(), unit_title.strip(), model.strip(), secret("OPENAI_API_KEY"))
                st.session_state["prepared_note"] = prepared.note
                st.session_state["source_files"] = [(f.name, f.type, f.getvalue()) for f in uploads]
                st.session_state["page_range"] = page_range
                st.session_state["vocab_analysis"] = result.model_dump()
                st.session_state.pop("test_bundle", None)
                st.session_state.pop("exports", None)
                st.success(f"어휘 후보 {len(result.vocabulary)}개를 선별했습니다.")
        except Exception as e:
            st.exception(e)

if "vocab_analysis" in st.session_state:
    analysis = core.VocabularyAnalysis.model_validate(st.session_state["vocab_analysis"])
    st.divider()
    st.subheader("2️⃣ Vocabulary List 검토")
    st.caption(st.session_state.get("prepared_note", ""))
    st.write(analysis.source_summary_ko)

    df = pd.DataFrame([
        {
            "사용": True,
            "Word": v.word,
            "IPA": v.pronunciation,
            "품사": v.pos,
            "뜻": v.meaning_ko,
            "Syn./Ant.": v.syn_ant,
            "New Example": v.new_example,
            "근거": v.source_evidence,
        }
        for v in analysis.vocabulary
    ])
    edited = st.data_editor(
        df,
        use_container_width=True,
        hide_index=True,
        num_rows="dynamic",
        column_config={"사용": st.column_config.CheckboxColumn(default=True)},
        key="vocab_editor",
    )

    def editor_to_vocab(frame: pd.DataFrame) -> list[core.VocabularyItem]:
        out = []
        for _, row in frame.iterrows():
            if not bool(row.get("사용", True)):
                continue
            out.append(core.VocabularyItem(
                word=str(row.get("Word", "")).strip(),
                pronunciation=str(row.get("IPA", "")).strip(),
                pos=str(row.get("품사", "")).strip(),
                meaning_ko=str(row.get("뜻", "")).strip(),
                syn_ant=str(row.get("Syn./Ant.", "—")).strip() or "—",
                new_example=str(row.get("New Example", "")).strip(),
                source_evidence=str(row.get("근거", "")).strip(),
            ))
        return out

    try:
        approved_vocab = editor_to_vocab(edited)
        vissues = core.validate_vocab(approved_vocab)
    except Exception as e:
        approved_vocab = []
        vissues = [str(e)]

    c1, c2 = st.columns([1, 2])
    with c1:
        st.metric("승인 어휘 수", len(approved_vocab))
    with c2:
        if vissues:
            st.error("어휘 검수: " + " / ".join(vissues[:5]))
        else:
            st.success("어휘 기본 검수 통과")

    if st.button("3️⃣ 이 어휘로 시험·답안 생성", use_container_width=True, disabled=bool(vissues)):
        try:
            saved = st.session_state["source_files"]

            class MemoryUpload:
                def __init__(self, name, typ, data):
                    self.name, self.type, self._data = name, typ, data
                def getvalue(self):
                    return self._data

            restored = [MemoryUpload(*x) for x in saved]
            prepared = core.prepare_sources(restored, st.session_state.get("page_range", ""))
            with st.spinner("Reading Review 12 + Transfer 9 + Definition 5 + Relations 4를 생성하고 있습니다..."):
                bundle = core.generate_test(prepared, approved_vocab, title.strip(), unit_title.strip(), model.strip(), secret("OPENAI_API_KEY"))
                st.session_state["approved_vocab"] = [v.model_dump() for v in approved_vocab]
                st.session_state["test_bundle"] = bundle.model_dump()
                st.session_state.pop("exports", None)
                st.rerun()
        except Exception as e:
            st.exception(e)

if "test_bundle" in st.session_state:
    st.divider()
    st.subheader("4️⃣ 시험·답안 검수")
    vocab = [core.VocabularyItem.model_validate(x) for x in st.session_state["approved_vocab"]]
    bundle = core.TestBundle.model_validate(st.session_state["test_bundle"])
    issues = core.validate_vocab(vocab) + core.validate_test(bundle, vocab)

    if issues:
        st.error("자동 검수에서 문제가 발견되었습니다.")
        for issue in issues:
            st.write("-", issue)
        st.info("시험을 다시 생성하거나 Vocabulary List를 수정한 뒤 다시 생성해 주세요.")
    else:
        st.success("자동 검수 통과: 30문항 구조·정답 대응·승인 어휘 사용·New Example 재사용 검사를 통과했습니다.")

    tabs = st.tabs(["시험지 미리보기", "답안·해설", "출력"])
    with tabs[0]:
        current_section = None
        for q in sorted(bundle.questions, key=lambda x: x.number):
            if q.section != current_section:
                current_section = q.section
                names = {1:"1. Reading Review",2:"2. Vocabulary Transfer",3:"3. English Definition",4:"4. Vocabulary Relations"}
                st.markdown(f"### {names[q.section]}")
            prompt, _ = core.render_question(q)
            st.markdown(f"**{q.number}.** {prompt}")
    with tabs[1]:
        rows=[]
        for q in sorted(bundle.questions, key=lambda x: x.number):
            _, answer = core.render_question(q)
            rows.append({"No.":q.number,"정답":answer,"뜻":q.meaning_ko,"문제 해설":q.explanation_ko})
        st.dataframe(rows, use_container_width=True, hide_index=True)
    with tabs[2]:
        if not issues:
            if st.button("5️⃣ Word + PDF 최종 파일 만들기", type="primary", use_container_width=True):
                with st.spinner("고정 템플릿에 내용을 넣고 Word/PDF를 생성합니다..."):
                    files, export_issues = core.build_export_bundle(title.strip(), unit_title.strip(), vocab, bundle)
                    st.session_state["exports"] = files
                    st.session_state["export_issues"] = export_issues
                    st.rerun()

        if "exports" in st.session_state:
            export_issues = st.session_state.get("export_issues", [])
            if export_issues:
                st.warning("Word는 생성되었지만 일부 출력 단계에 경고가 있습니다: " + " / ".join(export_issues))
            files = st.session_state["exports"]
            order = [
                "Vocabulary_List.docx","Vocabulary_List.pdf",
                "Vocabulary_Test.docx","Vocabulary_Test.pdf",
                "Answer_Explanation.docx","Answer_Explanation.pdf",
                "All_Files.zip",
            ]
            for name in order:
                if name not in files:
                    continue
                mime = "application/zip" if name.endswith(".zip") else "application/pdf" if name.endswith(".pdf") else "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                st.download_button(f"⬇️ {name}", files[name], file_name=name, mime=mime, use_container_width=True)

with st.sidebar:
    st.markdown("### 배포 체크")
    st.write("OPENAI_API_KEY:", "✅" if secret("OPENAI_API_KEY") else "❌")
    st.write("LibreOffice:", "✅" if (shutil.which("libreoffice") or shutil.which("soffice")) else "❌")
    st.caption("Secrets에는 API 키만 저장하고, 코드에 키를 직접 넣지 마세요.")
