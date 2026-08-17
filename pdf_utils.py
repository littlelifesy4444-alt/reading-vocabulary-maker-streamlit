# -*- coding: utf-8 -*-
"""
pdf_utils.py
PDF 업로드 파일에서 실제 텍스트를 추출하는 유틸리티.
Reading Vocabulary Maker MASTER MANUAL v1.0 - 3.1/3.2 절 기준:
  - 독해 자료는 PDF를 기본 입력으로 한다.
  - PDF가 길 경우 필요한 페이지 범위만 지정하여 분석한다.
"""

import io
from typing import Optional

import pdfplumber


class PdfExtractionError(Exception):
    """PDF 파싱/추출 중 발생한 오류를 감싸는 예외."""
    pass


def get_pdf_page_count(file_bytes: bytes) -> int:
    """업로드된 PDF 바이트에서 전체 페이지 수를 반환한다."""
    try:
        with pdfplumber.open(io.BytesIO(file_bytes)) as pdf:
            return len(pdf.pages)
    except Exception as e:
        raise PdfExtractionError(f"PDF 페이지 수를 읽는 중 오류가 발생했습니다: {e}")


def extract_text(file_bytes: bytes, start_page: int = 1, end_page: Optional[int] = None):
    """
    PDF 바이트에서 start_page~end_page(1-indexed, 포함) 범위의 텍스트를 추출한다.
    반환값: (추출된 텍스트, 전체 페이지 수, 실제로 텍스트가 추출된 페이지 수)
    """
    try:
        with pdfplumber.open(io.BytesIO(file_bytes)) as pdf:
            total_pages = len(pdf.pages)
            start_idx = max(1, int(start_page)) - 1
            end_idx = total_pages if not end_page else min(total_pages, int(end_page))
            if start_idx >= end_idx:
                raise PdfExtractionError(
                    f"페이지 범위가 올바르지 않습니다. (시작 {start_page}, 끝 {end_page}, 전체 {total_pages}페이지)"
                )

            page_texts = []
            pages_with_text = 0
            for i in range(start_idx, end_idx):
                page = pdf.pages[i]
                raw = page.extract_text() or ""
                raw = raw.strip()
                if raw:
                    pages_with_text += 1
                    page_texts.append(f"[Page {i + 1}]\n{raw}")

            full_text = "\n\n".join(page_texts).strip()

            if not full_text:
                raise PdfExtractionError(
                    "선택한 페이지 범위에서 텍스트를 추출하지 못했습니다. "
                    "스캔 이미지로만 구성된 PDF일 수 있습니다. 텍스트 기반 PDF를 업로드해주세요."
                )

            return full_text, total_pages, pages_with_text
    except PdfExtractionError:
        raise
    except Exception as e:
        raise PdfExtractionError(f"PDF 텍스트 추출 중 오류가 발생했습니다: {e}")


def truncate_text_for_ai(text: str, max_chars: int = 50000):
    """
    텍스트가 너무 길 경우 AI 호출용으로 길이를 제한한다.
    반환값: (사용할 텍스트, 잘렸는지 여부)
    """
    if len(text) <= max_chars:
        return text, False
    return text[:max_chars], True
