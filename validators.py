# -*- coding: utf-8 -*-
"""
validators.py
Reading Vocabulary Maker MASTER MANUAL v1.0 - 8장 "자동 검수 체크리스트"를
그대로 구현한 검증 로직.

체크리스트:
  1. 중복 어휘가 없는가?
  2. 필수 필드가 비어 있지 않은가?
  3. 발음과 품사가 존재하는가?
  4. 시험이 정확히 30문항인가?
  5. 문항 번호와 유형 배치가 1-12/13-21/22-26/27-30으로 정확한가?
  6. 1-21번 target word가 승인된 Vocabulary List에 존재하는가?
  7. 22-30번 객관식에 각각 4개의 선택지가 있는가?
  8. 객관식 문제마다 정답이 하나만 존재하는가?
  9. Vocabulary List의 New Example이 시험 문장에 그대로 재사용되지 않았는가?
  10. 시험지와 정답지의 1-30번이 정확히 대응하는가?
  (11. 최종 Word가 세로 방향인가? -> docx_builder.py 가 항상 portrait로 고정 생성하므로 여기서는 구조 검증만 수행)
"""

TYPE_RANGES = {
    "reading_review": range(1, 13),
    "vocabulary_transfer": range(13, 22),
    "english_definition": range(22, 27),
    "vocabulary_relations": range(27, 31),
}
TYPE_EXPECTED_COUNT = {
    "reading_review": 12,
    "vocabulary_transfer": 9,
    "english_definition": 5,
    "vocabulary_relations": 4,
}
MC_TYPES = ("english_definition", "vocabulary_relations")

VOCAB_REQUIRED_FIELDS = ["word", "pronunciation", "pos", "meaning", "new_example", "source"]


def validate_vocab_list(vocab_list):
    """Vocabulary List에 대한 검수. 문제 목록(list[str])을 반환한다. 빈 리스트면 통과."""
    issues = []

    if not vocab_list:
        issues.append("Vocabulary List가 비어 있습니다. 최소 1개 이상의 어휘가 필요합니다.")
        return issues

    seen = {}
    for idx, item in enumerate(vocab_list, start=1):
        word = str(item.get("word", "")).strip()
        label = word if word else f"{idx}번째 항목"

        if not word:
            issues.append(f"[{idx}번째 항목] Word/Expression이 비어 있습니다.")
        else:
            key = word.lower()
            if key in seen:
                issues.append(f"중복 어휘: '{word}' ({seen[key]}번째 항목과 {idx}번째 항목)")
            else:
                seen[key] = idx

        for field in VOCAB_REQUIRED_FIELDS:
            value = str(item.get(field, "")).strip()
            if not value:
                issues.append(f"'{label}': 필수 필드 '{field}' 값이 비어 있습니다.")

        pronunciation = str(item.get("pronunciation", "")).strip()
        pos = str(item.get("pos", "")).strip()
        if not pronunciation:
            issues.append(f"'{label}': 발음(pronunciation) 정보가 없습니다.")
        if not pos:
            issues.append(f"'{label}': 품사(part of speech) 정보가 없습니다.")

        importance = item.get("importance", None)
        try:
            imp_val = int(importance)
            if imp_val < 1 or imp_val > 5:
                raise ValueError()
        except (TypeError, ValueError):
            issues.append(f"'{label}': Importance 값은 1~5 사이의 정수여야 합니다 (현재: {importance!r}).")

    return issues


def validate_test(questions, vocab_list):
    """Vocabulary Review Test(30문항)에 대한 검수. 문제 목록(list[str])을 반환한다."""
    issues = []

    approved_words = {str(v.get("word", "")).strip().lower() for v in vocab_list if str(v.get("word", "")).strip()}

    # 4. 정확히 30문항인가
    if len(questions) != 30:
        issues.append(f"시험 문항 수가 30개가 아닙니다 (현재 {len(questions)}개).")

    seen_nos = set()
    nums_by_type = {}

    for q in questions:
        no = q.get("no")
        qtype = q.get("type")

        if no is None:
            issues.append("문항 번호(no)가 누락된 문항이 있습니다.")
            continue

        if no in seen_nos:
            issues.append(f"{no}번 문항 번호가 중복되었습니다.")
        seen_nos.add(no)
        nums_by_type.setdefault(qtype, []).append(no)

        # 5. 유형 배치가 규정된 번호 범위와 일치하는가
        expected_range = TYPE_RANGES.get(qtype)
        if expected_range is None:
            issues.append(f"{no}번 문항의 유형 '{qtype}'은(는) 허용되지 않는 유형입니다.")
        elif no not in expected_range:
            issues.append(f"{no}번 문항의 유형({qtype})이 규정된 번호 범위와 맞지 않습니다.")

        # 6. 1-21번 target word가 승인된 Vocabulary List에 존재하는가
        if isinstance(no, int) and no <= 21:
            target = str(q.get("target_word", "")).strip().lower()
            if not target or target not in approved_words:
                issues.append(
                    f"{no}번 문항의 target word '{q.get('target_word', '')}'가 승인된 Vocabulary List에 없습니다."
                )

        # 7 & 8. 22-30번 객관식: 선택지 4개 + 정답 하나만 존재
        if qtype in MC_TYPES:
            choices = q.get("choices") or []
            if len(choices) != 4:
                issues.append(f"{no}번 문항(4지선다)의 선택지가 4개가 아닙니다 (현재 {len(choices)}개).")
            answer_index = q.get("answer_index")
            if answer_index is None or not isinstance(answer_index, int) or not (0 <= answer_index < len(choices)):
                issues.append(f"{no}번 문항의 정답 인덱스(answer_index)가 올바르지 않습니다.")
            if choices:
                normalized_choices = [c.strip().lower() for c in choices if isinstance(c, str)]
                if len(normalized_choices) != len(set(normalized_choices)):
                    issues.append(f"{no}번 문항의 선택지 중 중복된 내용이 있습니다.")
        else:
            if not str(q.get("answer", "")).strip():
                issues.append(f"{no}번 문항의 정답(answer)이 비어 있습니다.")
            question_text = str(q.get("question_text", ""))
            if "___" not in question_text:
                issues.append(f"{no}번 문항에 빈칸('___' 형식)이 없습니다.")

        if not str(q.get("explanation", "")).strip():
            issues.append(f"{no}번 문항에 해설(explanation)이 없습니다.")

    # 유형별 문항 수 검증
    for qtype, expected_count in TYPE_EXPECTED_COUNT.items():
        actual = len(nums_by_type.get(qtype, []))
        if actual != expected_count:
            label = {
                "reading_review": "Reading Review",
                "vocabulary_transfer": "Vocabulary Transfer",
                "english_definition": "English Definition",
                "vocabulary_relations": "Vocabulary Relations",
            }[qtype]
            issues.append(f"{label} 유형 문항 수가 {expected_count}개가 아닙니다 (현재 {actual}개).")

    # 10. 1-30번이 정확히 대응하는가 (번호 1~30이 각각 정확히 한 번씩 존재)
    expected_nos = set(range(1, 31))
    if seen_nos != expected_nos:
        missing = sorted(expected_nos - seen_nos)
        extra = sorted(seen_nos - expected_nos)
        if missing:
            issues.append(f"시험지에 없는 문항 번호: {missing}")
        if extra:
            issues.append(f"허용 범위(1~30) 밖의 문항 번호: {extra}")

    # 9. New Example 재사용 검사
    examples = {}
    for v in vocab_list:
        word = str(v.get("word", "")).strip().lower()
        example = str(v.get("new_example", "")).strip()
        if word and example and len(example) >= 8:
            examples[word] = example

    for q in questions:
        qt = str(q.get("question_text", "")).strip().lower()
        target = str(q.get("target_word", "")).strip().lower()
        example = examples.get(target)
        if example and example.lower() in qt:
            issues.append(
                f"{q.get('no')}번 문항이 Vocabulary List의 New Example 문장을 그대로 재사용했습니다 (target: '{q.get('target_word')}')."
            )

    # 같은 목표 어휘의 과도한 반복 확인 (참고용 경고)
    from collections import Counter
    target_counter = Counter(
        str(q.get("target_word", "")).strip().lower() for q in questions if q.get("target_word")
    )
    for word, count in target_counter.items():
        if count > 2:
            issues.append(f"목표 어휘 '{word}'가 시험에서 {count}회 반복 사용되었습니다. 과도한 반복은 피해야 합니다.")

    return issues
