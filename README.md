# Reading Vocabulary Maker

Reading Vocabulary Maker MASTER MANUAL v1.0 규격을 그대로 구현한 Streamlit 앱입니다.

PDF 업로드 → 실제 PDF 텍스트 분석 → Vocabulary List 생성(AI) → 검토·수정·삭제·추가
→ 승인 → 30문항 Vocabulary Review Test 생성(AI) → 자동 검수 → Vocabulary List / Test /
Answer & Explanation Word 파일 생성 → ZIP 다운로드까지 전 과정이 실제로 동작합니다.

## 1. 설치

Python 3.10~3.13 권장. 배포 환경에서는 지원되는 안정 버전을 사용하세요.

```bash
pip install -r requirements.txt
```

## 2. 실행

```bash
streamlit run app.py
```

브라우저(또는 태블릿 Chrome)에서 안내된 주소(기본 `http://localhost:8501`)로 접속합니다.

## 3. 사용 준비물

- **Anthropic API Key** — https://console.anthropic.com 에서 발급. 앱 왼쪽 사이드바에 입력합니다.
  (키는 서버/파일에 저장되지 않고 현재 세션에서만 사용됩니다.)
- **텍스트 기반 PDF** — 스캔 이미지로만 이루어진 PDF는 텍스트 추출이 되지 않습니다.

## 4. 사용 절차 (MASTER MANUAL 10장 기준)

1. 교재명/Chapter명/시험지 제목을 입력하고 난이도를 선택합니다.
2. 독해 PDF를 업로드하고 필요하면 페이지 범위를 지정합니다.
3. "PDF 분석 및 Vocabulary List 생성"을 실행합니다.
4. 생성된 Vocabulary List 표를 직접 검토·수정·삭제·추가합니다.
5. "검수 실행"으로 문제 여부를 확인한 뒤 "최종 어휘 승인"을 누릅니다.
6. "30문항 시험 생성"을 실행합니다. (Reading Review 1-12 / Vocabulary Transfer 13-21 /
   English Definition 22-26 / Vocabulary Relations 27-30)
7. 자동 검수 체크리스트를 확인합니다. 문제가 있으면 "검수 오류 반영하여 재생성"으로
   다시 생성할 수 있습니다.
8. 검수를 모두 통과하면 "Word 문서 3종 생성"을 눌러 ZIP으로 다운로드합니다.
9. 최종 파일을 열어 페이지 잘림, 표 깨짐, 번호 오류 등을 육안으로 마지막 확인 후
   수업에 사용합니다.

## 5. 파일 구성

| 파일 | 역할 |
|---|---|
| `app.py` | Streamlit UI 및 전체 워크플로우 오케스트레이션 |
| `pdf_utils.py` | PDF 업로드 파일에서 실제 텍스트를 추출 |
| `ai_engine.py` | Claude API 호출 — Vocabulary List / 30문항 시험 생성 (JSON) |
| `validators.py` | MASTER MANUAL 8장 자동 검수 체크리스트 구현 |
| `docx_builder.py` | MASTER MANUAL 9장 디자인 잠금 규칙에 따른 Word 문서 3종 생성 |
| `requirements.txt` | 의존성 목록 |

## 6. 설계 원칙 (MASTER MANUAL 1장)

AI는 구조화된 내용 데이터(JSON)만 생성하고, 문서의 고정 디자인(세로 방향, 폰트,
여백, 제목 위치, footer "Hard work pays off.", 표 구조/테두리, No. 열 폭 등)은
`docx_builder.py`의 Python 코드가 담당합니다. 내용 생성과 문서 디자인이 분리되어
있으므로 AI 모델이 바뀌어도 최종 문서 레이아웃은 항상 동일하게 유지됩니다.

## 7. 태블릿(Chrome) 파일 업로드 관련

Samsung Android 태블릿 Chrome 등 모바일 브라우저 호환성을 위해 `st.file_uploader`를
`st.form` 없이 최상위 위젯으로 단순하게 구성했습니다. 다중 업로드나 커스텀 드래그
앤 드롭 없이 표준 파일 선택 방식만 사용합니다.


## 배포 권장 설정

- 기본 모델: `claude-sonnet-5`
- 선택 모델: `claude-opus-5`, `claude-haiku-4-5-20251001`, `claude-fable-5`
- 긴 PDF는 한 번에 너무 많은 페이지를 선택하기보다 단원/지문 단위로 범위를 나누어 분석하는 것을 권장합니다.
