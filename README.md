# Reading Vocabulary Maker v2

독해 PDF/이미지를 올리면 **Vocabulary List → 30문항 Vocabulary Review Test → Answer & Explanation**을 생성하는 Streamlit 앱입니다.

이 버전의 핵심은 **AI가 문서 디자인을 만들지 않는 것**입니다. AI는 어휘/문항/해설의 구조화된 데이터만 만들고, 최종 Word는 `templates/`의 고정 Word 템플릿에 내용만 삽입합니다.

## 배포 순서 — Streamlit Community Cloud

1. 이 폴더 전체를 GitHub 저장소에 올립니다.
2. Streamlit Community Cloud에서 새 앱을 만들고 **Main file path**를 `streamlit_app.py`로 지정합니다.
3. App settings → Secrets에 아래를 등록합니다.

```toml
OPENAI_API_KEY = "sk-..."
OPENAI_MODEL = "gpt-5.6"
```

4. Deploy 합니다.

`packages.txt`가 LibreOffice와 Noto CJK 폰트를 설치하도록 설정되어 있습니다. LibreOffice는 Word → PDF 변환에 사용됩니다.

## 사용법

1. 교재/Chapter 제목, 단원명을 입력합니다.
2. PDF 또는 이미지를 업로드합니다.
3. PDF 1개인 경우 필요하면 **PDF 내부 페이지 범위**를 입력합니다. 예: `35-38`.
4. **어휘 분석**을 누릅니다.
5. 화면에서 어휘를 삭제/수정/추가합니다. 어휘 수는 50개로 고정하지 않습니다.
6. **이 어휘로 시험·답안 생성**을 누릅니다.
7. 자동 검수를 확인합니다.
8. 검수 통과 후 **Word + PDF 최종 파일 만들기**를 누릅니다.
9. 6개 파일 또는 ZIP을 다운로드합니다.

## 고정 시험 구조

- 1. Reading Review: 1–12
- 2. Vocabulary Transfer: 13–21
- 3. English Definition: 22–26
- 4. Vocabulary Relations: 27–30

## 왜 이전 버전보다 안정적인가

- Vocabulary List를 먼저 사람이 승인한 뒤 시험을 생성합니다.
- OpenAI Structured Outputs(Pydantic)를 사용해 데이터 구조를 고정합니다.
- 문서 서식은 Chapter 6 최종 Word 파일을 템플릿으로 사용합니다.
- 문서 출력 전에 문항 수, 섹션, target word, 객관식 정답, New Example 재사용을 검사합니다.
- 페이지 방향을 코드로 검사하며 가로 문서가 나오면 출력 단계에서 중단합니다.

## 폴더 구조

```text
Reading_Vocabulary_Maker_v2/
├─ streamlit_app.py  ← Streamlit Cloud main file
├─ app.py            ← 호환용 진입점
├─ core.py           ← 생성/검수/Word-PDF 엔진
├─ requirements.txt
├─ packages.txt
├─ MANUAL.md
├─ README.md
├─ .streamlit/
│  ├─ config.toml
│  └─ secrets.toml.example
└─ templates/
   ├─ Vocabulary_List_Template.docx
   ├─ Vocabulary_Test_Template.docx
   └─ Answer_Explanation_Template.docx
```

## 주의

- API 사용량에 따라 OpenAI API 비용이 발생할 수 있습니다.
- 스캔 PDF는 시각 입력으로 처리되므로 페이지 수가 많으면 비용과 시간이 커질 수 있습니다. 필요한 페이지 범위를 잘라 사용하는 것을 권장합니다.
- `OPENAI_API_KEY`를 GitHub에 커밋하지 마세요.
