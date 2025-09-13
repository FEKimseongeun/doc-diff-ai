import os
import json
import difflib
from typing import Dict, List, Tuple, Optional, Any
from datetime import datetime
from .extractors import ExcelExtractor, WordExtractor, PDFExtractor


class DocumentComparator:
    """문서 비교 핵심 엔진"""

    def __init__(self):
        self.excel_extractor = ExcelExtractor()
        self.word_extractor = WordExtractor()
        self.pdf_extractor = PDFExtractor()

    def compare(self, original_path: str, revised_path: str,
                options: Optional[Dict] = None) -> Dict:
        """
        두 문서를 비교하여 변경사항 추출

        Returns:
            {
                "summary": {...},
                "changes": [...],
                "details": {...}
            }
        """
        options = options or {}

        # 파일 타입 감지
        original_type = self._detect_file_type(original_path)
        revised_type = self._detect_file_type(revised_path)

        if original_type != revised_type:
            raise ValueError("원본과 수정본의 파일 형식이 다릅니다.")

        # 문서 내용 추출
        original_content = self._extract_content(original_path, original_type)
        revised_content = self._extract_content(revised_path, revised_type)

        # 비교 실행
        if original_type == "excel":
            changes = self._compare_excel(original_content, revised_content, options)
        elif original_type == "word":
            changes = self._compare_word(original_content, revised_content, options)
        elif original_type == "pdf":
            changes = self._compare_pdf(original_content, revised_content, options)
        else:
            raise ValueError(f"지원하지 않는 파일 형식: {original_type}")

        # 결과 정리
        results = {
            "metadata": {
                "original_file": os.path.basename(original_path),
                "revised_file": os.path.basename(revised_path),
                "file_type": original_type,
                "compared_at": datetime.now().isoformat(),
                "options": options
            },
            "summary": self._generate_summary(changes),
            "changes": changes,
            "details": {
                "original_content": original_content,
                "revised_content": revised_content
            }
        }

        return results

    def _detect_file_type(self, file_path: str) -> str:
        """파일 확장자로 타입 감지"""
        ext = os.path.splitext(file_path)[1].lower()
        if ext in ['.xlsx', '.xls']:
            return 'excel'
        elif ext in ['.docx', '.doc']:
            return 'word'
        elif ext == '.pdf':
            return 'pdf'
        else:
            raise ValueError(f"지원하지 않는 파일 형식: {ext}")

    def _extract_content(self, file_path: str, file_type: str) -> Any:
        """파일 타입에 따라 내용 추출"""
        if file_type == 'excel':
            return self.excel_extractor.extract(file_path)
        elif file_type == 'word':
            return self.word_extractor.extract(file_path)
        elif file_type == 'pdf':
            return self.pdf_extractor.extract(file_path)

    def _compare_excel(self, original: Dict, revised: Dict,
                       options: Dict) -> List[Dict]:
        """Excel 파일 비교"""
        changes = []

        # 시트별 비교
        all_sheets = set(original.keys()) | set(revised.keys())

        for sheet in all_sheets:
            if sheet not in original:
                # 새로 추가된 시트
                changes.append({
                    "type": "sheet_added",
                    "sheet": sheet,
                    "content": revised[sheet]
                })
            elif sheet not in revised:
                # 삭제된 시트
                changes.append({
                    "type": "sheet_deleted",
                    "sheet": sheet,
                    "content": original[sheet]
                })
            else:
                # 셀 단위 비교
                sheet_changes = self._compare_excel_cells(
                    original[sheet], revised[sheet], sheet, options
                )
                changes.extend(sheet_changes)

        return changes

    def _compare_excel_cells(self, original_sheet: List[List],
                             revised_sheet: List[List],
                             sheet_name: str, options: Dict) -> List[Dict]:
        """Excel 셀 단위 상세 비교"""
        changes = []
        max_rows = max(len(original_sheet), len(revised_sheet))
        max_cols = max(
            max(len(row) for row in original_sheet) if original_sheet else 0,
            max(len(row) for row in revised_sheet) if revised_sheet else 0
        )

        for row_idx in range(max_rows):
            for col_idx in range(max_cols):
                original_val = self._get_cell_value(original_sheet, row_idx, col_idx)
                revised_val = self._get_cell_value(revised_sheet, row_idx, col_idx)

                if original_val != revised_val:
                    # 옵션 적용
                    if options.get('ignore_case'):
                        if str(original_val).lower() == str(revised_val).lower():
                            continue

                    if options.get('ignore_whitespace'):
                        if str(original_val).strip() == str(revised_val).strip():
                            continue

                    change = {
                        "type": "cell_modified",
                        "sheet": sheet_name,
                        "location": {
                            "row": row_idx + 1,  # 1-based indexing
                            "column": self._col_num_to_letter(col_idx + 1),
                            "cell": f"{self._col_num_to_letter(col_idx + 1)}{row_idx + 1}"
                        },
                        "original": original_val,
                        "revised": revised_val,
                        "change_type": self._classify_change(original_val, revised_val)
                    }

                    # 컨텍스트 추가
                    if options.get('show_context'):
                        change['context'] = self._get_cell_context(
                            original_sheet, revised_sheet,
                            row_idx, col_idx,
                            options.get('context_lines', 2)
                        )

                    changes.append(change)

        return changes

    def _compare_word(self, original: List[str], revised: List[str],
                      options: Dict) -> List[Dict]:
        """Word 문서 비교 - 문장/단락 단위"""
        changes = []

        # 문장 단위 비교를 위한 전처리
        original_sentences = self._split_sentences(original)
        revised_sentences = self._split_sentences(revised)

        # difflib를 사용한 상세 비교
        matcher = difflib.SequenceMatcher(None, original_sentences, revised_sentences)

        for tag, i1, i2, j1, j2 in matcher.get_opcodes():
            if tag == 'replace':
                for idx in range(i2 - i1):
                    if i1 + idx < len(original_sentences) and j1 + idx < len(revised_sentences):
                        changes.append({
                            "type": "sentence_modified",
                            "location": {
                                "original_index": i1 + idx,
                                "revised_index": j1 + idx
                            },
                            "original": original_sentences[i1 + idx],
                            "revised": revised_sentences[j1 + idx],
                            "similarity": difflib.SequenceMatcher(
                                None,
                                original_sentences[i1 + idx],
                                revised_sentences[j1 + idx]
                            ).ratio()
                        })
            elif tag == 'delete':
                for idx in range(i1, i2):
                    changes.append({
                        "type": "sentence_deleted",
                        "location": {"index": idx},
                        "original": original_sentences[idx],
                        "revised": None
                    })
            elif tag == 'insert':
                for idx in range(j1, j2):
                    changes.append({
                        "type": "sentence_added",
                        "location": {"index": idx},
                        "original": None,
                        "revised": revised_sentences[idx]
                    })

        return changes

    def _compare_pdf(self, original: Dict, revised: Dict,
                     options: Dict) -> List[Dict]:
        """PDF 문서 비교 - 페이지/텍스트 단위"""
        changes = []

        # 페이지별 비교
        max_pages = max(len(original['pages']), len(revised['pages']))

        for page_num in range(max_pages):
            if page_num >= len(original['pages']):
                # 새 페이지 추가됨
                changes.append({
                    "type": "page_added",
                    "page": page_num + 1,
                    "content": revised['pages'][page_num]
                })
            elif page_num >= len(revised['pages']):
                # 페이지 삭제됨
                changes.append({
                    "type": "page_deleted",
                    "page": page_num + 1,
                    "content": original['pages'][page_num]
                })
            else:
                # 페이지 내용 비교
                page_changes = self._compare_pdf_page(
                    original['pages'][page_num],
                    revised['pages'][page_num],
                    page_num + 1,
                    options
                )
                changes.extend(page_changes)

        return changes

    def _compare_pdf_page(self, original_page: str, revised_page: str,
                          page_num: int, options: Dict) -> List[Dict]:
        """PDF 페이지 내용 상세 비교"""
        changes = []

        # 라인 단위 비교
        original_lines = original_page.split('\n')
        revised_lines = revised_page.split('\n')

        matcher = difflib.SequenceMatcher(None, original_lines, revised_lines)

        for tag, i1, i2, j1, j2 in matcher.get_opcodes():
            if tag != 'equal':
                change = {
                    "type": f"pdf_line_{tag}",
                    "page": page_num,
                    "location": {
                        "original_lines": (i1 + 1, i2),
                        "revised_lines": (j1 + 1, j2)
                    }
                }

                if tag == 'replace':
                    change["original"] = '\n'.join(original_lines[i1:i2])
                    change["revised"] = '\n'.join(revised_lines[j1:j2])
                elif tag == 'delete':
                    change["original"] = '\n'.join(original_lines[i1:i2])
                    change["revised"] = None
                elif tag == 'insert':
                    change["original"] = None
                    change["revised"] = '\n'.join(revised_lines[j1:j2])

                changes.append(change)

        return changes

    # 유틸리티 메서드들
    def _get_cell_value(self, sheet: List[List], row: int, col: int):
        """안전하게 셀 값 가져오기"""
        if row < len(sheet) and col < len(sheet[row]):
            return sheet[row][col]
        return None

    def _col_num_to_letter(self, col_num: int) -> str:
        """열 번호를 Excel 열 문자로 변환 (1 -> A, 27 -> AA)"""
        result = ""
        while col_num > 0:
            col_num -= 1
            result = chr(65 + col_num % 26) + result
            col_num //= 26
        return result

    def _classify_change(self, original, revised) -> str:
        """변경 유형 분류"""
        if original is None or original == "":
            return "added"
        elif revised is None or revised == "":
            return "deleted"
        else:
            return "modified"

    def _get_cell_context(self, original_sheet: List[List],
                          revised_sheet: List[List],
                          row: int, col: int, context_size: int) -> Dict:
        """셀 주변 컨텍스트 가져오기"""
        context = {
            "original_surrounding": [],
            "revised_surrounding": []
        }

        for r in range(max(0, row - context_size),
                       min(len(original_sheet), row + context_size + 1)):
            for c in range(max(0, col - context_size),
                           min(len(original_sheet[r]) if r < len(original_sheet) else 0,
                               col + context_size + 1)):
                if r != row or c != col:
                    context["original_surrounding"].append({
                        "cell": f"{self._col_num_to_letter(c + 1)}{r + 1}",
                        "value": self._get_cell_value(original_sheet, r, c)
                    })

        for r in range(max(0, row - context_size),
                       min(len(revised_sheet), row + context_size + 1)):
            for c in range(max(0, col - context_size),
                           min(len(revised_sheet[r]) if r < len(revised_sheet) else 0,
                               col + context_size + 1)):
                if r != row or c != col:
                    context["revised_surrounding"].append({
                        "cell": f"{self._col_num_to_letter(c + 1)}{r + 1}",
                        "value": self._get_cell_value(revised_sheet, r, c)
                    })

        return context

    def _split_sentences(self, paragraphs: List[str]) -> List[str]:
        """단락을 문장으로 분리"""
        sentences = []
        for para in paragraphs:
            # 간단한 문장 분리 (향후 개선 가능)
            import re
            sents = re.split(r'[.!?]+', para)
            sentences.extend([s.strip() for s in sents if s.strip()])
        return sentences

    def _generate_summary(self, changes: List[Dict]) -> Dict:
        """변경사항 요약 생성"""
        summary = {
            "total_changes": len(changes),
            "changes_by_type": {},
            "severity": "low"  # low, medium, high
        }

        for change in changes:
            change_type = change["type"]
            if change_type not in summary["changes_by_type"]:
                summary["changes_by_type"][change_type] = 0
            summary["changes_by_type"][change_type] += 1

        # 심각도 판단
        if len(changes) > 50:
            summary["severity"] = "high"
        elif len(changes) > 10:
            summary["severity"] = "medium"

        return summary

    def save_html_report(self, results: Dict, output_path: str):
        """HTML 리포트 생성"""
        from .report_generator import HTMLReportGenerator
        generator = HTMLReportGenerator()
        html = generator.generate(results)
        with open(output_path, 'w', encoding='utf-8') as f:
            f.write(html)

    def save_json_results(self, results: Dict, output_path: str):
        """JSON 형식으로 결과 저장"""
        with open(output_path, 'w', encoding='utf-8') as f:
            json.dump(results, f, ensure_ascii=False, indent=2)