# Reading Vocabulary Maker — Streamlit 전용

이 버전은 처음부터 Streamlit Community Cloud 배포를 전제로 만든 새 버전입니다.

## 파일
- `streamlit_app.py` — 실행 파일
- `requirements.txt` — 필요한 Python 패키지

## Streamlit 배포값
- Repository: 새 저장소
- Branch: `main`
- Main file path: `streamlit_app.py`

## Secrets
Streamlit 앱의 **Settings → Secrets**에 아래 형식으로 입력합니다.

```toml
OPENAI_API_KEY = "본인의_API_키"
```

모델을 따로 지정하고 싶을 때만 추가:

```toml
OPENAI_MODEL = "사용할_모델명"
```

## 앱 기능
- PDF/이미지 업로드
- PDF 페이지 범위 지정
- 난이도: 더 쉽게 / 기본 / 더 어렵게
- 지문 난이도·어휘 밀도에 따른 어휘 수 자동 조절
- 단어장 생성
- 4유형 시험 생성
- 정답·해설 생성
- 문제별 쉽게/어렵게 재생성
- 시험 순서 재섞기
- 단어장 / 시험지 / 정답해설 Word 다운로드

## 시험 매뉴얼
1. 지문 복습형 주관식 + 첫 글자
2. 새 문맥 주관식 + 첫 글자
3. 영영풀이 객관식
4. 정확한 동의어·반의어·어휘관계 객관식

단어장 새 예문은 시험에 그대로 재사용하지 않습니다.
정확한 동의어·반의어가 없으면 `—`로 표시합니다.
