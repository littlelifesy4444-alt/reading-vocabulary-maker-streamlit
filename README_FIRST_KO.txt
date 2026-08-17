Reading Vocabulary Maker v2 — 먼저 읽어주세요

1) 이 폴더를 GitHub 저장소에 그대로 올립니다.
2) Streamlit Community Cloud에서 새 앱을 만듭니다.
3) Repository/Branch는 해당 저장소의 main을 선택합니다.
4) Main file path는 streamlit_app.py 입니다.
5) Settings → Secrets에 아래 두 줄을 넣습니다.

OPENAI_API_KEY = "본인의 API 키"
OPENAI_MODEL = "gpt-5.6"

6) Deploy 합니다.

이번 버전은 이전 버전과 달리:
- AI가 서식을 만들지 않습니다.
- Chapter 6 최종 Word 3종을 고정 템플릿으로 씁니다.
- 어휘를 먼저 검토/승인한 뒤 시험을 만듭니다.
- 어휘 수를 50개로 강제하지 않습니다.
- 시험 30문항 구조를 자동 검수합니다.
- New Example을 시험에 그대로 재사용했는지 검사합니다.
- 세로 페이지 여부를 검사합니다.
- Word와 PDF를 함께 만듭니다.

큰 PDF가 50MB를 넘더라도 업로드는 가능하지만 OpenAI로 보낼 때는 50MB 미만이어야 합니다.
따라서 큰 교재 PDF는 반드시 'PDF 내부 페이지 범위'를 지정해 필요한 Chapter만 잘라 분석하세요.
