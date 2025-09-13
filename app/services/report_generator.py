# app/services/report_generator.py
from typing import Dict, List, Any
import html


class HTMLReportGenerator:
    """HTML 형식의 비교 리포트 생성기"""

    def generate(self, results: Dict) -> str:
        """비교 결과를 HTML로 변환"""

        html_content = f"""
<!DOCTYPE html>
<html lang="ko">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>문서 비교 결과</title>
    <style>
        {self._get_css()}
    </style>
</head>
<body>
    <div class="container">
        <header>
            <h1>📄 문서 비교 결과</h1>
            <div class="metadata">
                <div class="meta-item">
                    <span class="label">원본 파일:</span>
                    <span class="value">{results['metadata']['original_file']}</span>
                </div>
                <div class="meta-item">
                    <span class="label">수정본 파일:</span>
                    <span class="value">{results['metadata']['revised_file']}</span>
                </div>
                <div class="meta-item">
                    <span class="label">파일 형식:</span>
                    <span class="value">{results['metadata']['file_type'].upper()}</span>
                </div>
                <div class="meta-item">
                    <span class="label">비교 시간:</span>
                    <span class="value">{results['metadata']['compared_at']}</span>
                </div>
            </div>
        </header>

        <section class="summary">
            <h2>📊 요약</h2>
            <div class="summary-grid">
                <div class="summary-card">
                    <div class="summary-number">{results['summary']['total_changes']}</div>
                    <div class="summary-label">전체 변경사항</div>
                </div>
                <div class="summary-card severity-{results['summary']['severity']}">
                    <div class="summary-number">{results['summary']['severity'].upper()}</div>
                    <div class="summary-label">변경 수준</div>
                </div>
            </div>

            <div class="change-types">
                <h3>변경 유형별 통계</h3>
                <div class="type-grid">
                    {self._generate_type_stats(results['summary']['changes_by_type'])}
                </div>
            </div>
        </section>

        <section class="changes">
            <h2>🔍 상세 변경사항</h2>
            <div class="filter-controls">
                <button onclick="filterChanges('all')" class="filter-btn active">전체</button>
                <button onclick="filterChanges('modified')" class="filter-btn">수정</button>
                <button onclick="filterChanges('added')" class="filter-btn">추가</button>
                <button onclick="filterChanges('deleted')" class="filter-btn">삭제</button>
            </div>

            <div class="changes-list">
                {self._generate_changes_html(results['changes'], results['metadata']['file_type'])}
            </div>
        </section>
    </div>

    <script>
        {self._get_javascript()}
    </script>
</body>
</html>
"""
        return html_content

    def _get_css(self) -> str:
        """CSS 스타일 정의"""
        return """
        * {
            margin: 0;
            padding: 0;
            box-sizing: border-box;
        }

        body {
            font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, Oxygen, Ubuntu, sans-serif;
            background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
            min-height: 100vh;
            padding: 20px;
        }

        .container {
            max-width: 1200px;
            margin: 0 auto;
            background: white;
            border-radius: 20px;
            box-shadow: 0 20px 60px rgba(0,0,0,0.3);
            overflow: hidden;
        }

        header {
            background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
            color: white;
            padding: 30px;
        }

        header h1 {
            font-size: 2.5rem;
            margin-bottom: 20px;
        }

        .metadata {
            display: grid;
            grid-template-columns: repeat(auto-fit, minmax(250px, 1fr));
            gap: 15px;
        }

        .meta-item {
            background: rgba(255,255,255,0.1);
            padding: 10px 15px;
            border-radius: 10px;
            backdrop-filter: blur(10px);
        }

        .meta-item .label {
            font-weight: 600;
            margin-right: 10px;
        }

        .meta-item .value {
            color: #ffd700;
        }

        .summary {
            padding: 30px;
            background: #f8f9fa;
        }

        .summary h2 {
            color: #333;
            margin-bottom: 20px;
            font-size: 1.8rem;
        }

        .summary-grid {
            display: grid;
            grid-template-columns: repeat(auto-fit, minmax(200px, 1fr));
            gap: 20px;
            margin-bottom: 30px;
        }

        .summary-card {
            background: white;
            padding: 20px;
            border-radius: 15px;
            text-align: center;
            box-shadow: 0 4px 6px rgba(0,0,0,0.1);
            transition: transform 0.3s;
        }

        .summary-card:hover {
            transform: translateY(-5px);
        }

        .summary-number {
            font-size: 2.5rem;
            font-weight: bold;
            color: #667eea;
            margin-bottom: 10px;
        }

        .severity-low .summary-number { color: #28a745; }
        .severity-medium .summary-number { color: #ffc107; }
        .severity-high .summary-number { color: #dc3545; }

        .change-types {
            background: white;
            padding: 20px;
            border-radius: 15px;
        }

        .type-grid {
            display: grid;
            grid-template-columns: repeat(auto-fill, minmax(150px, 1fr));
            gap: 10px;
        }

        .type-stat {
            background: #f1f3f5;
            padding: 10px;
            border-radius: 8px;
            display: flex;
            justify-content: space-between;
            align-items: center;
        }

        .type-stat .count {
            font-weight: bold;
            color: #667eea;
        }

        .changes {
            padding: 30px;
        }

        .changes h2 {
            color: #333;
            margin-bottom: 20px;
            font-size: 1.8rem;
        }

        .filter-controls {
            margin-bottom: 20px;
            display: flex;
            gap: 10px;
        }

        .filter-btn {
            padding: 8px 20px;
            border: none;
            background: #e9ecef;
            border-radius: 20px;
            cursor: pointer;
            transition: all 0.3s;
            font-weight: 500;
        }

        .filter-btn:hover {
            background: #dee2e6;
        }

        .filter-btn.active {
            background: #667eea;
            color: white;
        }

        .changes-list {
            display: grid;
            gap: 15px;
        }

        .change-item {
            background: white;
            border: 1px solid #e9ecef;
            border-radius: 12px;
            padding: 20px;
            transition: all 0.3s;
        }

        .change-item:hover {
            box-shadow: 0 5px 15px rgba(0,0,0,0.1);
        }

        .change-header {
            display: flex;
            justify-content: space-between;
            align-items: center;
            margin-bottom: 15px;
        }

        .change-type {
            padding: 4px 12px;
            border-radius: 20px;
            font-size: 0.85rem;
            font-weight: 600;
            color: white;
        }

        .type-modified { background: #ffc107; }
        .type-added { background: #28a745; }
        .type-deleted { background: #dc3545; }

        .change-location {
            color: #6c757d;
            font-size: 0.9rem;
        }

        .change-content {
            display: grid;
            gap: 10px;
        }

        .content-block {
            padding: 15px;
            border-radius: 8px;
            background: #f8f9fa;
            position: relative;
        }

        .content-label {
            position: absolute;
            top: -8px;
            left: 10px;
            background: white;
            padding: 2px 8px;
            border-radius: 10px;
            font-size: 0.75rem;
            font-weight: 600;
            color: #6c757d;
        }

        .original-content {
            background: #fff5f5;
            border: 1px solid #ffcccc;
        }

        .revised-content {
            background: #f5fff5;
            border: 1px solid #ccffcc;
        }

        .content-text {
            margin-top: 8px;
            word-break: break-word;
            font-family: 'Courier New', monospace;
        }

        .null-value {
            color: #999;
            font-style: italic;
        }

        .context-info {
            margin-top: 10px;
            padding: 10px;
            background: #f1f3f5;
            border-radius: 8px;
            font-size: 0.85rem;
        }

        .no-changes {
            text-align: center;
            padding: 60px;
            color: #6c757d;
        }

        .no-changes svg {
            width: 100px;
            height: 100px;
            margin-bottom: 20px;
            opacity: 0.5;
        }
        """

    def _get_javascript(self) -> str:
        """JavaScript 코드"""
        return """
        function filterChanges(type) {
            // 필터 버튼 활성화 상태 변경
            document.querySelectorAll('.filter-btn').forEach(btn => {
                btn.classList.remove('active');
            });
            event.target.classList.add('active');

            // 변경사항 필터링
            const items = document.querySelectorAll('.change-item');
            items.forEach(item => {
                if (type === 'all') {
                    item.style.display = 'block';
                } else {
                    const hasType = item.classList.contains('filter-' + type);
                    item.style.display = hasType ? 'block' : 'none';
                }
            });
        }

        // 변경사항 펼치기/접기
        document.addEventListener('DOMContentLoaded', function() {
            document.querySelectorAll('.change-header').forEach(header => {
                header.style.cursor = 'pointer';
                header.addEventListener('click', function() {
                    const content = this.nextElementSibling;
                    content.style.display = content.style.display === 'none' ? 'grid' : 'none';
                });
            });
        });
        """

    def _generate_type_stats(self, changes_by_type: Dict) -> str:
        """변경 유형별 통계 HTML 생성"""
        html_parts = []
        for change_type, count in changes_by_type.items():
            type_label = self._format_type_label(change_type)
            html_parts.append(f"""
                <div class="type-stat">
                    <span>{type_label}</span>
                    <span class="count">{count}</span>
                </div>
            """)
        return ''.join(html_parts)

    def _generate_changes_html(self, changes: List[Dict], file_type: str) -> str:
        """변경사항 상세 HTML 생성"""
        if not changes:
            return """
                <div class="no-changes">
                    <svg viewBox="0 0 24 24" fill="currentColor">
                        <path d="M9 11H7v2h2v-2zm4 0h-2v2h2v-2zm4 0h-2v2h2v-2zm2-7h-1V2h-2v2H8V2H6v2H5c-1.11 0-1.99.9-1.99 2L3 20c0 1.1.89 2 2 2h14c1.1 0 2-.9 2-2V6c0-1.1-.9-2-2-2zm0 16H5V9h14v11z"/>
                    </svg>
                    <h3>변경사항이 없습니다</h3>
                    <p>두 문서가 동일합니다.</p>
                </div>
            """

        html_parts = []
        for idx, change in enumerate(changes):
            change_class = self._get_change_class(change)
            filter_class = self._get_filter_class(change)

            html_parts.append(f"""
                <div class="change-item {filter_class}" data-index="{idx}">
                    <div class="change-header">
                        <div>
                            <span class="change-type type-{change_class}">{self._format_type_label(change['type'])}</span>
                            <span class="change-location">{self._format_location(change, file_type)}</span>
                        </div>
                    </div>
                    <div class="change-content">
                        {self._format_change_content(change)}
                    </div>
                </div>
            """)

        return ''.join(html_parts)

    def _format_change_content(self, change: Dict) -> str:
        """변경 내용 포맷팅"""
        content_html = []

        # 원본 내용
        if 'original' in change:
            original_text = self._escape_and_format(change['original'])
            content_html.append(f"""
                <div class="content-block original-content">
                    <span class="content-label">원본</span>
                    <div class="content-text">{original_text}</div>
                </div>
            """)

        # 수정본 내용
        if 'revised' in change:
            revised_text = self._escape_and_format(change['revised'])
            content_html.append(f"""
                <div class="content-block revised-content">
                    <span class="content-label">수정본</span>
                    <div class="content-text">{revised_text}</div>
                </div>
            """)

        # 컨텍스트 정보
        if 'context' in change:
            content_html.append(self._format_context(change['context']))

        return ''.join(content_html)

    def _escape_and_format(self, value: Any) -> str:
        """값 이스케이프 및 포맷팅"""
        if value is None:
            return '<span class="null-value">(없음)</span>'

        text = str(value)
        # HTML 이스케이프
        text = html.escape(text)
        # 줄바꿈 처리
        text = text.replace('\n', '<br>')
        return text

    def _format_location(self, change: Dict, file_type: str) -> str:
        """위치 정보 포맷팅"""
        location = change.get('location', {})

        if file_type == 'excel':
            if 'cell' in location:
                return f"셀 {location['cell']}"
            elif 'sheet' in change:
                return f"시트: {change['sheet']}"
        elif file_type == 'word':
            if 'original_index' in location:
                return f"문장 #{location['original_index'] + 1}"
        elif file_type == 'pdf':
            if 'page' in change:
                return f"페이지 {change['page']}"

        return ""

    def _format_context(self, context: Dict) -> str:
        """컨텍스트 정보 포맷팅"""
        html_parts = ['<div class="context-info"><strong>주변 컨텍스트:</strong><br>']

        if 'original_surrounding' in context:
            html_parts.append('<em>원본 주변 셀:</em> ')
            cells = [f"{c['cell']}: {c['value']}" for c in context['original_surrounding'][:3]]
            html_parts.append(', '.join(cells))

        if 'revised_surrounding' in context:
            html_parts.append('<br><em>수정본 주변 셀:</em> ')
            cells = [f"{c['cell']}: {c['value']}" for c in context['revised_surrounding'][:3]]
            html_parts.append(', '.join(cells))

        html_parts.append('</div>')
        return ''.join(html_parts)

    def _get_change_class(self, change: Dict) -> str:
        """변경 유형에 따른 CSS 클래스"""
        change_type = change.get('change_type', '')
        if 'added' in change_type or 'add' in change['type']:
            return 'added'
        elif 'deleted' in change_type or 'delete' in change['type']:
            return 'deleted'
        else:
            return 'modified'

    def _get_filter_class(self, change: Dict) -> str:
        """필터용 CSS 클래스"""
        change_class = self._get_change_class(change)
        return f"filter-{change_class}"

    def _format_type_label(self, type_str: str) -> str:
        """타입 문자열을 읽기 쉬운 형태로 변환"""
        type_map = {
            'cell_modified': '셀 수정',
            'sentence_modified': '문장 수정',
            'sentence_added': '문장 추가',
            'sentence_deleted': '문장 삭제',
            'sheet_added': '시트 추가',
            'sheet_deleted': '시트 삭제',
            'page_added': '페이지 추가',
            'page_deleted': '페이지 삭제',
            'pdf_line_replace': 'PDF 라인 수정',
            'pdf_line_insert': 'PDF 라인 추가',
            'pdf_line_delete': 'PDF 라인 삭제'
        }
        return type_map.get(type_str, type_str.replace('_', ' ').title())