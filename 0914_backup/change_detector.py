# change_detector.py
import base64
import html
import io
import re
from difflib import SequenceMatcher
from typing import Dict, List, Any

import numpy as np
from PIL import Image


# ========= Diff Utilities =========
def _tokenize_keep_spaces(s: str) -> List[str]:
    return re.findall(r'\w+|[^\w\s]+|\s+', s or "", re.UNICODE)

def _word_diff_html(old: str, new: str):
    a = _tokenize_keep_spaces(old or "")
    b = _tokenize_keep_spaces(new or "")
    sm = SequenceMatcher(a=a, b=b)
    pieces = []
    added, deleted = [], []
    for tag, i1, i2, j1, j2 in sm.get_opcodes():
        if tag == 'equal':
            pieces.append(''.join(html.escape(x) for x in a[i1:i2]))
        elif tag == 'insert':
            seg = ''.join(html.escape(x) for x in b[j1:j2])
            pieces.append(f'<ins class="diff-add">{seg}</ins>')
            added += [t for t in b[j1:j2] if t.strip() and not t.isspace()]
        elif tag == 'delete':
            seg = ''.join(html.escape(x) for x in a[i1:i2])
            pieces.append(f'<del class="diff-del">{seg}</del>')
            deleted += [t for t in a[i1:i2] if t.strip() and not t.isspace()]
        elif tag == 'replace':
            seg_old = ''.join(html.escape(x) for x in a[i1:i2])
            seg_new = ''.join(html.escape(x) for x in b[j1:j2])
            pieces.append(f'<del class="diff-del">{seg_old}</del><ins class="diff-add">{seg_new}</ins>')
            added   += [t for t in b[j1:j2] if t.strip() and not t.isspace()]
            deleted += [t for t in a[i1:i2] if t.strip() and not t.isspace()]
    return ''.join(pieces), added, deleted

def _split_sentences(s: str) -> List[str]:
    parts = re.split(r'(?<=[\.\!\?])\s+|\n+', s or "")
    return [p for p in parts if p]


# ========= Change Detector =========
class ChangeDetector:
    def __init__(self):
        self.similarity_threshold = 0.8
        self.image_similarity_threshold = 0.95

    def detect_changes(self, original: Dict[str, Any], revised: Dict[str, Any]) -> Dict[str, Any]:
        changes = {
            'text_changes': [],
            'formatting_changes': [],
            'table_changes': [],
            'image_changes': [],
            'annotation_changes': [],   # ← 추가
            'structural_changes': [],
            'summary': {
                'total_changes': 0,
                'text_changes_count': 0,
                'formatting_changes_count': 0,
                'table_changes_count': 0,
                'image_changes_count': 0,
                'annotation_changes_count': 0,   # ← 추가
                'structural_changes_count': 0,
            }
        }

        # Text
        text_changes = self._detect_text_changes(original, revised)
        changes['text_changes'] = text_changes
        changes['summary']['text_changes_count'] = len(text_changes)

        # Formatting
        formatting_changes = self._detect_formatting_changes(original, revised)
        changes['formatting_changes'] = formatting_changes
        changes['summary']['formatting_changes_count'] = len(formatting_changes)

        # Tables
        table_changes = self._detect_table_changes(original, revised)
        changes['table_changes'] = table_changes
        changes['summary']['table_changes_count'] = len(table_changes)

        # Images
        image_changes = self._detect_image_changes(original, revised)
        changes['image_changes'] = image_changes
        changes['summary']['image_changes_count'] = len(image_changes)

        # Annotations (PDF)
        annotation_changes = self._detect_annotation_changes(original, revised)
        changes['annotation_changes'] = annotation_changes
        changes['summary']['annotation_changes_count'] = len(annotation_changes)

        # Structural
        structural_changes = self._detect_structural_changes(original, revised)
        changes['structural_changes'] = structural_changes
        changes['summary']['structural_changes_count'] = len(structural_changes)

        # Total
        changes['summary']['total_changes'] = sum([
            changes['summary']['text_changes_count'],
            changes['summary']['formatting_changes_count'],
            changes['summary']['table_changes_count'],
            changes['summary']['image_changes_count'],
            changes['summary']['annotation_changes_count'],
            changes['summary']['structural_changes_count'],
        ])
        return changes

    # ---------- TEXT ----------
    def _detect_text_changes(self, original: Dict[str, Any], revised: Dict[str, Any]) -> List[Dict[str, Any]]:
        out = []
        if original['type'] == 'docx' and revised['type'] == 'docx':
            out.extend(self._detect_docx_text_changes(original, revised))
        elif original['type'] == 'pdf' and revised['type'] == 'pdf':
            out.extend(self._detect_pdf_text_changes(original, revised))
        elif original['type'] == 'xlsx' and revised['type'] == 'xlsx':
            out.extend(self._detect_xlsx_text_changes(original, revised))
        return out

    def _detect_docx_text_changes(self, original: Dict[str, Any], revised: Dict[str, Any]) -> List[Dict[str, Any]]:
        changes = []
        op = original.get('paragraphs', [])
        rp = revised.get('paragraphs', [])
        n = max(len(op), len(rp))
        for i in range(n):
            ot = op[i]['text'] if i < len(op) else ""
            nt = rp[i]['text'] if i < len(rp) else ""
            if ot and not nt:
                changes.append({'type': 'text_change','change_type':'deleted',
                                'content': ot,'paragraph_index': i,'document_type':'docx'})
            elif nt and not ot:
                changes.append({'type': 'text_change','change_type':'added',
                                'content': nt,'paragraph_index': i,'document_type':'docx'})
            elif ot != nt:
                diff_html, added, deleted = _word_diff_html(ot, nt)
                changes.append({'type':'text_change','change_type':'modified',
                                'content_html': diff_html,'old_text': ot,'new_text': nt,
                                'added_terms': added,'deleted_terms': deleted,
                                'paragraph_index': i,'document_type':'docx'})
        return changes

    def _detect_pdf_text_changes(self, original: Dict[str, Any], revised: Dict[str, Any]) -> List[Dict[str, Any]]:
        changes = []
        op = original.get('pages', [])
        rp = revised.get('pages', [])
        m = max(len(op), len(rp))
        for i in range(m):
            ot = (op[i].get('text') if i < len(op) else "") or ""
            nt = (rp[i].get('text') if i < len(rp) else "") or ""
            a, b = _split_sentences(ot), _split_sentences(nt)
            sm = SequenceMatcher(a=a, b=b)
            for tag, i1, i2, j1, j2 in sm.get_opcodes():
                if tag == 'equal': continue
                elif tag == 'delete':
                    for sent in a[i1:i2]:
                        changes.append({'type':'text_change','change_type':'deleted',
                                        'content': sent,'page_number': i+1,'document_type':'pdf'})
                elif tag == 'insert':
                    for sent in b[j1:j2]:
                        changes.append({'type':'text_change','change_type':'added',
                                        'content': sent,'page_number': i+1,'document_type':'pdf'})
                elif tag == 'replace':
                    old_seg, new_seg = a[i1:i2], b[j1:j2]
                    k = min(len(old_seg), len(new_seg))
                    for s_old, s_new in zip(old_seg[:k], new_seg[:k]):
                        diff_html, added, deleted = _word_diff_html(s_old, s_new)
                        changes.append({'type':'text_change','change_type':'modified',
                                        'content_html': diff_html,'old_text': s_old,'new_text': s_new,
                                        'added_terms': added,'deleted_terms': deleted,
                                        'page_number': i+1,'document_type':'pdf'})
                    for sent in old_seg[k:]:
                        changes.append({'type':'text_change','change_type':'deleted',
                                        'content': sent,'page_number': i+1,'document_type':'pdf'})
                    for sent in new_seg[k:]:
                        changes.append({'type':'text_change','change_type':'added',
                                        'content': sent,'page_number': i+1,'document_type':'pdf'})
        return changes

    # ---------- XLSX (Text) ----------
    def _to_str(self, v) -> str:
        return "" if v is None else str(v)

    def _detect_xlsx_text_changes(self, original: Dict[str, Any], revised: Dict[str, Any]) -> List[Dict[str, Any]]:
        changes = []
        original_sheets = original.get('sheets', [])
        revised_sheets  = revised.get('sheets', [])
        sheet_names = {s['name'] for s in original_sheets} | {s['name'] for s in revised_sheets}
        for sheet_name in sheet_names:
            osheet = next((s for s in original_sheets if s['name'] == sheet_name), None)
            rsheet = next((s for s in revised_sheets  if s['name'] == sheet_name), None)
            if not osheet:
                changes.append({'type':'sheet_added','change_type':'added',
                                'content':f"Sheet added: {sheet_name}",'sheet_name':sheet_name,'document_type':'xlsx'})
                continue
            if not rsheet:
                changes.append({'type':'sheet_deleted','change_type':'deleted',
                                'content':f"Sheet deleted: {sheet_name}",'sheet_name':sheet_name,'document_type':'xlsx'})
                continue
            ocells = {c['coordinate']: c for row in osheet['cells'] for c in row}
            rcells = {c['coordinate']: c for row in rsheet['cells'] for c in row}
            allc = set(ocells)|set(rcells)
            for coord in allc:
                oc, rc = ocells.get(coord), rcells.get(coord)
                if not oc:
                    val = self._to_str(rc.get('value') if rc else "")
                    changes.append({'type':'cell_added','change_type':'added','content': val,'value':val,
                                    'coordinate':coord,'sheet_name':sheet_name,'document_type':'xlsx'})
                elif not rc:
                    val = self._to_str(oc.get('value'))
                    changes.append({'type':'cell_deleted','change_type':'deleted','content': val,'value':val,
                                    'coordinate':coord,'sheet_name':sheet_name,'document_type':'xlsx'})
                else:
                    old_val = self._to_str(oc.get('value'))
                    new_val = self._to_str(rc.get('value'))
                    if old_val != new_val:
                        diff_html, added, deleted = _word_diff_html(old_val, new_val)
                        changes.append({'type':'cell_modified','change_type':'modified',
                                        'coordinate':coord,'sheet_name':sheet_name,'document_type':'xlsx',
                                        'old_value': old_val,'new_value': new_val,
                                        'content_html': diff_html,'added_terms': added,'deleted_terms': deleted})
        return changes

    # ---------- ANNOTATIONS (PDF) ----------
    @staticmethod
    def _norm_floats(v):
        if v is None: return None
        try:
            return [round(float(x), 3) for x in v]
        except Exception:
            return None

    @staticmethod
    def _annot_key(a: Dict[str, Any]) -> str:
        """고유키: NM(id) 우선, 없으면 페이지/타입/Rect/내용으로 생성"""
        nm = a.get('id') or ""
        if nm: return nm
        page = a.get('page_number')
        subtype = a.get('subtype') or ""
        rect = ChangeDetector._norm_floats(a.get('rect')) or []
        contents = (a.get('contents') or "")[:64]
        base = f"{page}:{subtype}:{','.join(map(str,rect))}:{contents}"
        # 파서에서도 fallback id 생성하지만, 한 번 더 안전하게
        import hashlib
        return "AUTO-" + hashlib.md5(base.encode('utf-8')).hexdigest()[:10]

    def _detect_annotation_changes(self, original: Dict[str, Any], revised: Dict[str, Any]) -> List[Dict[str, Any]]:
        """PDF Annotation: 추가/삭제/수정"""
        if not (original['type'] == revised['type'] == 'pdf'):
            return []

        olist = original.get('annotations', []) or []
        rlist = revised.get('annotations', []) or []

        omap = {self._annot_key(a): a for a in olist}
        rmap = {self._annot_key(a): a for a in rlist}

        changes: List[Dict[str, Any]] = []

        # Added
        for k, a in rmap.items():
            if k not in omap:
                changes.append({
                    'type': 'annotation_change', 'change_type': 'added',
                    'document_type': 'pdf',
                    'id': a.get('id') or k,
                    'page_number': a.get('page_number'),
                    'subtype': a.get('subtype'),
                    'content': a.get('contents'),
                    'rect': a.get('rect'),
                    'quadpoints': a.get('quadpoints'),
                    'author': a.get('author'),
                    'subject': a.get('subject'),
                    'color': a.get('color'),
                })

        # Deleted
        for k, a in omap.items():
            if k not in rmap:
                changes.append({
                    'type': 'annotation_change', 'change_type': 'deleted',
                    'document_type': 'pdf',
                    'id': a.get('id') or k,
                    'page_number': a.get('page_number'),
                    'subtype': a.get('subtype'),
                    'content': a.get('contents'),
                    'rect': a.get('rect'),
                    'quadpoints': a.get('quadpoints'),
                    'author': a.get('author'),
                    'subject': a.get('subject'),
                    'color': a.get('color'),
                })

        # Modified
        for k in set(omap.keys()) & set(rmap.keys()):
            o = omap[k]; r = rmap[k]
            diffs = []

            # 내용 diff (텍스트 주석/프리텍스트/텍스트마크업에 유용)
            otext = o.get('contents') or ""
            rtext = r.get('contents') or ""
            diff_html = None
            if otext != rtext:
                diff_html, _, _ = _word_diff_html(otext, rtext)
                diffs.append("contents")

            # 위치/형태/색상 등 비교(소수점 3자리 반올림)
            if self._norm_floats(o.get('rect')) != self._norm_floats(r.get('rect')):
                diffs.append("rect")
            if self._norm_floats(o.get('quadpoints')) != self._norm_floats(r.get('quadpoints')):
                diffs.append("quadpoints")
            if self._norm_floats(o.get('color')) != self._norm_floats(r.get('color')):
                diffs.append("color")
            if (o.get('subtype') or "") != (r.get('subtype') or ""):
                diffs.append("subtype")
            if (o.get('author') or "") != (r.get('author') or ""):
                diffs.append("author")
            if (o.get('subject') or "") != (r.get('subject') or ""):
                diffs.append("subject")

            if diffs:
                changes.append({
                    'type': 'annotation_change', 'change_type': 'modified',
                    'document_type': 'pdf',
                    'id': r.get('id') or k,
                    'page_number': r.get('page_number'),
                    'subtype': r.get('subtype'),
                    'content_html': diff_html,        # 내용이 바뀐 경우 하이라이트
                    'old': {
                        'contents': otext, 'rect': o.get('rect'),
                        'quadpoints': o.get('quadpoints'), 'color': o.get('color'),
                        'author': o.get('author'), 'subject': o.get('subject'),
                        'subtype': o.get('subtype')
                    },
                    'new': {
                        'contents': rtext, 'rect': r.get('rect'),
                        'quadpoints': r.get('quadpoints'), 'color': r.get('color'),
                        'author': r.get('author'), 'subject': r.get('subject'),
                        'subtype': r.get('subtype')
                    },
                    'changed_fields': diffs
                })

        return changes

    # ---------- FORMATTING ----------
    def _detect_formatting_changes(self, original: Dict[str, Any], revised: Dict[str, Any]) -> List[Dict[str, Any]]:
        changes = []
        if original['type'] == 'docx' and revised['type'] == 'docx':
            changes.extend(self._detect_docx_formatting_changes(original, revised))
        elif original['type'] == 'xlsx' and revised['type'] == 'xlsx':
            changes.extend(self._detect_xlsx_formatting_changes(original, revised))
        return changes

    def _detect_docx_formatting_changes(self, original: Dict[str, Any], revised: Dict[str, Any]) -> List[Dict[str, Any]]:
        changes = []
        op = original.get('paragraphs', [])
        rp = revised.get('paragraphs', [])
        for i, (orig_para, rev_para) in enumerate(zip(op, rp)):
            if orig_para.get('style') != rev_para.get('style'):
                changes.append({'type':'paragraph_style_change','paragraph_index': i,
                                'old_style': orig_para.get('style'),'new_style': rev_para.get('style'),
                                'document_type':'docx'})
            oruns = orig_para.get('runs', [])
            rruns = rev_para.get('runs', [])
            for j, (orun, rrun) in enumerate(zip(oruns, rruns)):
                if orun.get('text') == rrun.get('text'):
                    fmt = self._compare_run_formatting(orun, rrun)
                    if fmt:
                        changes.append({'type':'run_formatting_change','paragraph_index': i,'run_index': j,
                                        'text': orun.get('text'),'changes': fmt,'document_type':'docx'})
        return changes

    def _compare_run_formatting(self, orig_run: Dict[str, Any], rev_run: Dict[str, Any]) -> List[str]:
        out = []
        for attr in ['bold','italic','underline','font_name','font_size','font_color']:
            if orig_run.get(attr) != rev_run.get(attr):
                out.append(f"{attr}: {orig_run.get(attr)} → {rev_run.get(attr)}")
        return out

    def _detect_xlsx_formatting_changes(self, original: Dict[str, Any], revised: Dict[str, Any]) -> List[Dict[str, Any]]:
        changes = []
        osheets = original.get('sheets', [])
        rsheets = revised.get('sheets', [])
        for os, rs in zip(osheets, rsheets):
            if os.get('name') != rs.get('name'): continue
            ocells = {c['coordinate']: c for row in os['cells'] for c in row}
            rcells = {c['coordinate']: c for row in rs['cells'] for c in row}
            for coord, oc in ocells.items():
                rc = rcells.get(coord)
                if not rc: continue
                if oc.get('value') == rc.get('value'):
                    fmt = self._compare_cell_formatting(oc, rc)
                    if fmt:
                        changes.append({'type':'cell_formatting_change','coordinate': coord,
                                        'sheet_name': os.get('name'),'changes': fmt,'document_type':'xlsx'})
        return changes

    def _compare_cell_formatting(self, oc: Dict[str, Any], rc: Dict[str, Any]) -> List[str]:
        out = []
        of, rf = oc.get('font', {}), rc.get('font', {})
        for attr in ['name','size','bold','italic','color']:
            if of.get(attr) != rf.get(attr):
                out.append(f"font_{attr}: {of.get(attr)} → {rf.get(attr)}")
        ofill, rfill = oc.get('fill', {}), rc.get('fill', {})
        if ofill.get('fg_color') != rfill.get('fg_color'):
            out.append(f"fill_color: {ofill.get('fg_color')} → {rfill.get('fg_color')}")
        ob, rb = oc.get('border', {}), rc.get('border', {})
        for s in ['left','right','top','bottom']:
            if ob.get(s) != rb.get(s):
                out.append(f"border_{s}: {ob.get(s)} → {rb.get(s)}")
        return out

    # ---------- TABLES ----------
    def _detect_table_changes(self, original: Dict[str, Any], revised: Dict[str, Any]) -> List[Dict[str, Any]]:
        changes = []
        if original['type'] == 'docx' and revised['type'] == 'docx':
            changes.extend(self._detect_docx_table_changes(original, revised))
        return changes

    def _detect_docx_table_changes(self, original: Dict[str, Any], revised: Dict[str, Any]) -> List[Dict[str, Any]]:
        changes = []
        otables = original.get('tables', [])
        rtables = revised.get('tables', [])
        if len(otables) != len(rtables):
            changes.append({'type':'table_count_change','original_count':len(otables),
                            'revised_count':len(rtables),'document_type':'docx'})
        for i, (ot, rt) in enumerate(zip(otables, rtables)):
            orows = ot.get('rows', []); rrows = rt.get('rows', [])
            if len(orows) != len(rrows):
                changes.append({'type':'table_row_count_change','table_index': i,
                                'original_rows': len(orows),'revised_rows': len(rrows),'document_type':'docx'})
            for j, (orow, rrow) in enumerate(zip(orows, rrows)):
                for k, (ocell, rcell) in enumerate(zip(orow, rrow)):
                    if ocell.get('text') != rcell.get('text'):
                        changes.append({'type':'table_cell_change','table_index': i,'row_index': j,'cell_index': k,
                                        'old_text': ocell.get('text'),'new_text': rcell.get('text'),'document_type':'docx'})
        return changes

    # ---------- IMAGES ----------
    def _detect_image_changes(self, original: Dict[str, Any], revised: Dict[str, Any]) -> List[Dict[str, Any]]:
        changes = []
        if original['type'] == 'docx' and revised['type'] == 'docx':
            changes.extend(self._detect_docx_image_changes(original, revised))
        return changes

    def _detect_docx_image_changes(self, original: Dict[str, Any], revised: Dict[str, Any]) -> List[Dict[str, Any]]:
        changes = []
        oimgs = original.get('images', []); rimgs = revised.get('images', [])
        if len(oimgs) != len(rimgs):
            changes.append({'type':'image_count_change','original_count':len(oimgs),
                            'revised_count':len(rimgs),'document_type':'docx'})
        for i, (oi, ri) in enumerate(zip(oimgs, rimgs)):
            if oi.get('size') != ri.get('size'):
                changes.append({'type':'image_size_change','image_index': i,
                                'old_size': oi.get('size'),'new_size': ri.get('size'),'document_type':'docx'})
            sim = self._compare_images(oi.get('data',''), ri.get('data',''))
            if sim < self.image_similarity_threshold:
                changes.append({'type':'image_content_change','image_index': i,
                                'similarity': sim,'document_type':'docx'})
        return changes

    def _compare_images(self, img1_data: str, img2_data: str) -> float:
        try:
            from skimage.metrics import structural_similarity as ssim
            img1 = Image.open(io.BytesIO(base64.b64decode(img1_data))).convert('RGB')
            img2 = Image.open(io.BytesIO(base64.b64decode(img2_data))).convert('RGB')
            a1, a2 = np.array(img1), np.array(img2)
            h, w = min(a1.shape[0], a2.shape[0]), min(a1.shape[1], a2.shape[1])
            a1, a2 = a1[:h, :w], a2[:h, :w]
            return float(ssim(a1, a2, channel_axis=-1))
        except Exception as e:
            print(f"Error comparing images: {e}")
            return 0.0

    # ---------- STRUCTURE ----------
    def _detect_structural_changes(self, original: Dict[str, Any], revised: Dict[str, Any]) -> List[Dict[str, Any]]:
        changes = []
        if original['type'] != revised['type']:
            changes.append({'type':'document_type_change','original_type': original['type'],
                            'revised_type': revised['type']})
            return changes
        if original['type'] == 'docx':
            oc, rc = len(original.get('paragraphs', [])), len(revised.get('paragraphs', []))
            if oc != rc:
                changes.append({'type':'paragraph_count_change','original_count': oc,'revised_count': rc,'document_type':'docx'})
        elif original['type'] == 'pdf':
            oc, rc = len(original.get('pages', [])), len(revised.get('pages', []))
            if oc != rc:
                changes.append({'type':'page_count_change','original_count': oc,'revised_count': rc,'document_type':'pdf'})
        elif original['type'] == 'xlsx':
            oc, rc = len(original.get('sheets', [])), len(revised.get('sheets', []))
            if oc != rc:
                changes.append({'type':'sheet_count_change','original_count': oc,'revised_count': rc,'document_type':'xlsx'})
        return changes
