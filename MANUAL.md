# Reading Vocabulary Maker — 고정 제작 매뉴얼

## 핵심 원칙
- AI는 **내용만 생성**한다. 문서 디자인은 Word 템플릿과 Python 코드가 고정한다.
- Vocabulary List의 어휘 수는 50개로 고정하지 않는다. 실제 지문에서 학습 가치가 있는 어휘만 선정한다.
- List의 New Example은 새 학습 문장으로 작성하고 시험 문장에 그대로 재사용하지 않는다.
- 시험은 항상 30문항: 1) Reading Review 1–12, 2) Vocabulary Transfer 13–21, 3) English Definition 22–26, 4) Vocabulary Relations 27–30.
- 1·2형은 첫 글자 힌트 + 빈칸형, 3·4형은 4지선다형.
- Answer & Explanation은 No. | 정답 | 뜻 | 문제 해설 4열이며, 이유를 설명한다.

## 디자인 잠금
- Chapter 6 최종본 Word 3종을 템플릿으로 사용한다.
- 세로 페이지, 글씨체, 글자 크기, 여백, 제목 위치, footer의 *Hard work pays off.*를 임의 변경하지 않는다.
- Vocabulary List의 모든 셀은 내부 가로/세로 포함 실선으로 구분한다.
- No. 칸은 좁게 유지하고 기존 템플릿 칸 너비를 유지한다.
- 시험지는 학생이 답을 쓰기 위한 행간/문항 간격을 유지한다.
- 답안지는 번호·정답 칸을 좁게, 해설 칸을 넓게 유지한다.

## 자동 검수
내보내기 전 다음을 모두 통과해야 한다.
1. 중복 어휘 없음 / 빈 필드 없음 / IPA+품사 존재.
2. 시험 정확히 30문항, 번호와 섹션 배치 정확.
3. 1–21번 target word가 승인된 Vocabulary List에 존재.
4. 22–30번은 4개 선택지와 하나의 정답만 존재.
5. List의 New Example이 시험 문장에 그대로 재사용되지 않음.
6. 시험지와 답안지 1–30번이 1:1 대응.
7. Word 생성 후 문서 페이지 방향이 세로인지 검사.
8. LibreOffice가 설치된 환경에서는 Word에서 PDF를 생성한다.
