import os
import json
from datetime import datetime
from typing import Dict, List, Any
import base64
from jinja2 import Template

class ReportGenerator:
    def __init__(self):
        self.reports_folder = 'reports'
        os.makedirs(self.reports_folder, exist_ok=True)
    
    def generate_report(self, changes: Dict[str, Any], original_content: Dict[str, Any], revised_content: Dict[str, Any]) -> str:
        """Generate a comprehensive comparison report."""
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        report_filename = f"comparison_report_{timestamp}.html"
        report_path = os.path.join(self.reports_folder, report_filename)
        
        # Generate HTML report
        html_content = self._generate_html_report(changes, original_content, revised_content)
        
        with open(report_path, 'w', encoding='utf-8') as f:
            f.write(html_content)
        
        return report_filename
    
    def _generate_html_report(self, changes: Dict[str, Any], original_content: Dict[str, Any], revised_content: Dict[str, Any]) -> str:
        """Generate HTML report with detailed change analysis."""
        
        html_template = """
<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Document Comparison Report</title>
    <style>
        body {
            font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif;
            line-height: 1.6;
            margin: 0;
            padding: 20px;
            background-color: #f5f5f5;
        }
        .container {
            max-width: 1200px;
            margin: 0 auto;
            background: white;
            padding: 30px;
            border-radius: 10px;
            box-shadow: 0 0 20px rgba(0,0,0,0.1);
        }
        .header {
            text-align: center;
            border-bottom: 3px solid #007acc;
            padding-bottom: 20px;
            margin-bottom: 30px;
        }
        .header h1 {
            color: #007acc;
            margin: 0;
            font-size: 2.5em;
        }
        .header p {
            color: #666;
            margin: 10px 0 0 0;
        }
        .summary {
            background: #f8f9fa;
            padding: 20px;
            border-radius: 8px;
            margin-bottom: 30px;
            border-left: 4px solid #007acc;
        }
        .summary h2 {
            margin-top: 0;
            color: #333;
        }
        .summary-grid {
            display: grid;
            grid-template-columns: repeat(auto-fit, minmax(200px, 1fr));
            gap: 15px;
            margin-top: 15px;
        }
        .summary-item {
            background: white;
            padding: 15px;
            border-radius: 5px;
            text-align: center;
            border: 1px solid #ddd;
        }
        .summary-item .number {
            font-size: 2em;
            font-weight: bold;
            color: #007acc;
        }
        .summary-item .label {
            color: #666;
            font-size: 0.9em;
        }
        .section {
            margin-bottom: 40px;
        }
        .section h2 {
            color: #333;
            border-bottom: 2px solid #007acc;
            padding-bottom: 10px;
        }
        .change-item {
            background: #f8f9fa;
            border: 1px solid #ddd;
            border-radius: 5px;
            padding: 15px;
            margin-bottom: 15px;
            position: relative;
        }
        .change-item.added {
            border-left: 4px solid #28a745;
            background: #f8fff9;
        }
        .change-item.deleted {
            border-left: 4px solid #dc3545;
            background: #fff8f8;
        }
        .change-item.modified {
            border-left: 4px solid #ffc107;
            background: #fffef8;
        }
        .change-type {
            display: inline-block;
            padding: 4px 8px;
            border-radius: 3px;
            font-size: 0.8em;
            font-weight: bold;
            text-transform: uppercase;
            margin-bottom: 10px;
        }
        .change-type.added {
            background: #28a745;
            color: white;
        }
        .change-type.deleted {
            background: #dc3545;
            color: white;
        }
        .change-type.modified {
            background: #ffc107;
            color: #333;
        }
        .change-content {
            font-family: 'Courier New', monospace;
            background: white;
            padding: 10px;
            border-radius: 3px;
            border: 1px solid #ddd;
            white-space: pre-wrap;
            word-break: break-word;
        }
        .change-details {
            margin-top: 10px;
            font-size: 0.9em;
            color: #666;
        }
        .image-comparison {
            display: flex;
            gap: 20px;
            margin-top: 15px;
        }
        .image-comparison img {
            max-width: 200px;
            max-height: 200px;
            border: 1px solid #ddd;
            border-radius: 5px;
        }
        .no-changes {
            text-align: center;
            color: #28a745;
            font-size: 1.2em;
            padding: 40px;
            background: #f8fff9;
            border-radius: 8px;
            border: 1px solid #28a745;
        }
        .footer {
            text-align: center;
            margin-top: 40px;
            padding-top: 20px;
            border-top: 1px solid #ddd;
            color: #666;
        }
        .table-comparison {
            overflow-x: auto;
        }
        .table-comparison table {
            width: 100%;
            border-collapse: collapse;
            margin-top: 10px;
        }
        .table-comparison th,
        .table-comparison td {
            border: 1px solid #ddd;
            padding: 8px;
            text-align: left;
        }
        .table-comparison th {
            background: #f8f9fa;
            font-weight: bold;
        }
        .cell-added {
            background: #d4edda;
        }
        .cell-deleted {
            background: #f8d7da;
        }
        .cell-modified {
            background: #fff3cd;
        }
          .diff-html ins.diff-add { background: #e8f5e9; text-decoration: none; border-radius: 3px; padding: 0 2px; }
          .diff-html del.diff-del { background: #ffebee; text-decoration: line-through; border-radius: 3px; padding: 0 2px; }
          .pill { display:inline-block; padding:2px 8px; border-radius:12px; font-size:12px; color:#fff; }
          .pill.added { background:#2e7d32; } .pill.deleted { background:#c62828; } .pill.modified { background:#1565c0; }
          .anno { border:1px solid #e0e0e0; padding:10px; border-radius:8px; margin:8px 0; }
          .anno-head { font-weight:600; margin-bottom:4px; }
          .meta { color:#666; font-size:.9em; }
    </style>
</head>
<body>
    <div class="container">
        <div class="header">
            <h1>Document Comparison Report</h1>
            <p>Generated on {{ timestamp }}</p>
        </div>
        
        <div class="summary">
            <h2>Summary</h2>
            <div class="summary-grid">
                <div class="summary-item">
                    <div class="number">{{ changes.summary.total_changes }}</div>
                    <div class="label">Total Changes</div>
                </div>
                <div class="summary-item">
                    <div class="number">{{ changes.summary.text_changes_count }}</div>
                    <div class="label">Text Changes</div>
                </div>
                <div class="summary-item">
                    <div class="number">{{ changes.summary.formatting_changes_count }}</div>
                    <div class="label">Formatting Changes</div>
                </div>
                <div class="summary-item">
                    <div class="number">{{ changes.summary.table_changes_count }}</div>
                    <div class="label">Table Changes</div>
                </div>
                <div class="summary-item">
                    <div class="number">{{ changes.summary.image_changes_count }}</div>
                    <div class="label">Image Changes</div>
                </div>
                <div class="summary-item">
                    <div class="number">{{ changes.summary.structural_changes_count }}</div>
                    <div class="label">Structural Changes</div>
                </div>
            </div>
        </div>
        
        {% if changes.text_changes %}
        <div class="section">
          <h2>Text Changes</h2>
          {% for change in changes.text_changes %}
          <div class="change-item {{ change.change_type }}">
            <div class="change-type {{ change.change_type }}">{{ change.change_type }}</div>
            <div class="change-content">
              {% if change.content_html %}
                <div class="diff-html">{{ change.content_html | safe }}</div>
              {% else %}
                {{ change.content }}
              {% endif %}
              {% if change.added_terms or change.deleted_terms %}
              <div class="diff-meta">
                {% if change.added_terms %}<strong>Added:</strong> {{ change.added_terms|join(' ') }}{% endif %}
                {% if change.deleted_terms %}{% if change.added_terms %} &nbsp;|&nbsp; {% endif %}<strong>Deleted:</strong> {{ change.deleted_terms|join(' ') }}{% endif %}
              </div>
              {% endif %}
            </div>
            <div class="change-details">
              Document Type: {{ change.document_type }}
              {% if change.page_number %} | Page: {{ change.page_number }}{% endif %}
              {% if change.paragraph_index is not none %} | Paragraph: {{ change.paragraph_index }}{% endif %}
              {% if change.sheet_name %} | Sheet: {{ change.sheet_name }}{% endif %}
              {% if change.coordinate %} | Cell: {{ change.coordinate }}{% endif %}
            </div>
          </div>
          {% endfor %}
        </div>
        {% endif %}
        
        {% if changes.formatting_changes %}
        <div class="section">
            <h2>Formatting Changes</h2>
            {% for change in changes.formatting_changes %}
            <div class="change-item modified">
                <div class="change-type modified">{{ change.type.replace('_', ' ').title() }}</div>
                {% if change.text %}
                <div class="change-content">{{ change.text }}</div>
                {% endif %}
                {% if change.changes %}
                <div class="change-details">
                    <strong>Changes:</strong>
                    <ul>
                        {% for format_change in change.changes %}
                        <li>{{ format_change }}</li>
                        {% endfor %}
                    </ul>
                </div>
                {% endif %}
                <div class="change-details">
                    Document Type: {{ change.document_type }}
                    {% if change.paragraph_index is defined %} | Paragraph: {{ change.paragraph_index }}{% endif %}
                    {% if change.run_index is defined %} | Run: {{ change.run_index }}{% endif %}
                    {% if change.coordinate %} | Cell: {{ change.coordinate }}{% endif %}
                    {% if change.sheet_name %} | Sheet: {{ change.sheet_name }}{% endif %}
                </div>
            </div>
            {% endfor %}
        </div>
        {% endif %}
        
        {% if changes.table_changes %}
        <div class="section">
            <h2>Table Changes</h2>
            {% for change in changes.table_changes %}
            <div class="change-item modified">
                <div class="change-type modified">{{ change.type.replace('_', ' ').title() }}</div>
                <div class="change-details">
                    {% if change.table_index is defined %}Table Index: {{ change.table_index }}<br>{% endif %}
                    {% if change.original_count is defined %}Original: {{ change.original_count }}<br>{% endif %}
                    {% if change.revised_count is defined %}Revised: {{ change.revised_count }}<br>{% endif %}
                    {% if change.original_rows is defined %}Original Rows: {{ change.original_rows }}<br>{% endif %}
                    {% if change.revised_rows is defined %}Revised Rows: {{ change.revised_rows }}<br>{% endif %}
                    {% if change.row_index is defined %}Row: {{ change.row_index }}<br>{% endif %}
                    {% if change.cell_index is defined %}Cell: {{ change.cell_index }}<br>{% endif %}
                    {% if change.old_text is defined %}
                    <strong>Old Text:</strong> {{ change.old_text }}<br>
                    {% endif %}
                    {% if change.new_text is defined %}
                    <strong>New Text:</strong> {{ change.new_text }}
                    {% endif %}
                </div>
            </div>
            {% endfor %}
        </div>
        {% endif %}
        
        {% if changes.annotation_changes %}
        <section>
          <h2>Annotation Changes</h2>
          {% for a in changes.annotation_changes %}
            <div class="anno">
              <div class="anno-head">
                <span class="pill {{ a.change_type }}">{{ a.change_type }}</span>
                <span>Subtype: {{ a.subtype or 'N/A' }}</span>
                <span class="meta">| Page {{ a.page_number }} | ID: {{ a.id }}</span>
              </div>
              <div>
                {% if a.content_html %}
                  <div class="diff-html">{{ a.content_html | safe }}</div>
                {% elif a.content %}
                  {{ a.content }}
                {% else %}
                  <em>(no text)</em>
                {% endif %}
              </div>
              <div class="meta">
                {% if a.changed_fields %}<strong>Changed:</strong> {{ a.changed_fields|join(', ') }}{% endif %}
                {% if a.rect %} | <strong>Rect:</strong> {{ a.rect }}{% endif %}
                {% if a.color %} | <strong>Color:</strong> {{ a.color }}{% endif %}
                {% if a.author %} | <strong>Author:</strong> {{ a.author }}{% endif %}
                {% if a.subject %} | <strong>Subject:</strong> {{ a.subject }}{% endif %}
              </div>
            </div>
          {% endfor %}
        </section>
        {% endif %}
        
        {% if changes.image_changes %}
        <div class="section">
            <h2>Image Changes</h2>
            {% for change in changes.image_changes %}
            <div class="change-item modified">
                <div class="change-type modified">{{ change.type.replace('_', ' ').title() }}</div>
                <div class="change-details">
                    {% if change.image_index is defined %}Image Index: {{ change.image_index }}<br>{% endif %}
                    {% if change.original_count is defined %}Original Count: {{ change.original_count }}<br>{% endif %}
                    {% if change.revised_count is defined %}Revised Count: {{ change.revised_count }}<br>{% endif %}
                    {% if change.old_size is defined %}Old Size: {{ change.old_size }}<br>{% endif %}
                    {% if change.new_size is defined %}New Size: {{ change.new_size }}<br>{% endif %}
                    {% if change.similarity is defined %}Similarity: {{ "%.2f"|format(change.similarity * 100) }}%{% endif %}
                </div>
            </div>
            {% endfor %}
        </div>
        {% endif %}
        
        {% if changes.structural_changes %}
        <div class="section">
            <h2>Structural Changes</h2>
            {% for change in changes.structural_changes %}
            <div class="change-item modified">
                <div class="change-type modified">{{ change.type.replace('_', ' ').title() }}</div>
                <div class="change-details">
                    {% if change.original_type is defined %}Original Type: {{ change.original_type }}<br>{% endif %}
                    {% if change.revised_type is defined %}Revised Type: {{ change.revised_type }}<br>{% endif %}
                    {% if change.original_count is defined %}Original Count: {{ change.original_count }}<br>{% endif %}
                    {% if change.revised_count is defined %}Revised Count: {{ change.revised_count }}<br>{% endif %}
                </div>
            </div>
            {% endfor %}
        </div>
        {% endif %}
        
        {% if changes.summary.total_changes == 0 %}
        <div class="no-changes">
            <h2>No Changes Detected</h2>
            <p>The documents appear to be identical.</p>
        </div>
        {% endif %}
        
        <div class="footer">
            <p>Report generated by Document Comparison Service</p>
        </div>
    </div>
</body>
</html>
        """
        
        template = Template(html_template)
        return template.render(
            changes=changes,
            original_content=original_content,
            revised_content=revised_content,
            timestamp=datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        )
    
    def generate_json_report(self, changes: Dict[str, Any], original_content: Dict[str, Any], revised_content: Dict[str, Any]) -> str:
        """Generate a JSON report for programmatic access."""
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        report_filename = f"comparison_report_{timestamp}.json"
        report_path = os.path.join(self.reports_folder, report_filename)
        
        report_data = {
            'timestamp': datetime.now().isoformat(),
            'changes': changes,
            'original_document_info': {
                'type': original_content.get('type'),
                'paragraph_count': len(original_content.get('paragraphs', [])),
                'table_count': len(original_content.get('tables', [])),
                'image_count': len(original_content.get('images', [])),
                'page_count': len(original_content.get('pages', [])),
                'sheet_count': len(original_content.get('sheets', []))
            },
            'revised_document_info': {
                'type': revised_content.get('type'),
                'paragraph_count': len(revised_content.get('paragraphs', [])),
                'table_count': len(revised_content.get('tables', [])),
                'image_count': len(revised_content.get('images', [])),
                'page_count': len(revised_content.get('pages', [])),
                'sheet_count': len(revised_content.get('sheets', []))
            }
        }
        
        with open(report_path, 'w', encoding='utf-8') as f:
            json.dump(report_data, f, indent=2, ensure_ascii=False)
        
        return report_filename
