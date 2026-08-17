# Reading Vocabulary Maker — OpenAI edition

MASTER MANUAL v1.0 기반 Streamlit 앱.

## Streamlit Secrets
앱 설정의 Secrets에 다음 이름으로 기존 OpenAI API 키를 등록합니다.

```toml
OPENAI_API_KEY = "..."
```

## 실행 파일
Streamlit Community Cloud의 기존 배포가 `streamlit_app.py`를 가리켜도 작동하도록
`app.py`와 동일한 `streamlit_app.py`를 함께 제공합니다.

## 주요 흐름
PDF 업로드 → 텍스트 추출 → OpenAI Vocabulary List 생성 → 교사 검토/승인 →
30문항 시험 생성 → 자동 검수 → Word 3종 → ZIP 다운로드.

## 파일
- streamlit_app.py / app.py: UI 및 전체 워크플로우
- ai_engine.py: OpenAI Responses API
- pdf_utils.py: PDF 텍스트 추출
- validators.py: 자동 검수
- docx_builder.py: Word 문서 생성
- requirements.txt: 의존성
