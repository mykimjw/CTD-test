import io
import os
import re
import copy
import difflib
import docx
from docx.shared import RGBColor, Pt, Inches
from docx.oxml import parse_xml, OxmlElement
from docx.oxml.ns import nsdecls, qn
from fastapi import FastAPI, UploadFile, File, Header, HTTPException
from fastapi.responses import StreamingResponse

app = FastAPI()

# 보안용 API 키 (Apps Script와 일치시켜야 함)
MY_API_KEY = os.getenv("MY_API_KEY", "my-secret-key-1234")

# 색상 상수 정의
COLOR_RED = RGBColor(255, 0, 0)
COLOR_BLUE = RGBColor(0, 0, 255)
COLOR_BLACK = RGBColor(0, 0, 0)
HEX_HEADER_BG = "E6EEF8"

# ----------------------------------------------------
# XML 및 서식 처리 헬퍼 함수들
# ----------------------------------------------------
def set_cell_background(cell, hex_color):
    tcPr = cell._tc.get_or_add_tcPr()
    shd = parse_xml(f'<w:shd {nsdecls("w")} w:fill="{hex_color}"/>')
    tcPr.append(shd)

def set_cell_margins(cell, top=100, bottom=100, left=150, right=150):
    tcPr = cell._tc.get_or_add_tcPr()
    tcMar = OxmlElement('w:tcMar')
    for m, val in [('w:top', top), ('w:bottom', bottom), ('w:left', left), ('w:right', right)]:
        node = OxmlElement(m)
        node.set(qn('w:w'), str(val))
        node.set(qn('w:type'), 'dxa')
        tcMar.append(node)
    tcPr.append(tcMar)

def set_table_borders(table):
    tblPr = table._tbl.tblPr
    borders = parse_xml(
        f'<w:tblBorders {nsdecls("w")}>'
        f'  <w:top w:val="single" w:sz="4" w:space="0" w:color="D3D3D3"/>'
        f'  <w:bottom w:val="single" w:sz="4" w:space="0" w:color="D3D3D3"/>'
        f'  <w:left w:val="single" w:sz="4" w:space="0" w:color="D3D3D3"/>'
        f'  <w:right w:val="single" w:sz="4" w:space="0" w:color="D3D3D3"/>'
        f'  <w:insideH w:val="single" w:sz="4" w:space="0" w:color="E0E0E0"/>'
        f'  <w:insideV w:val="single" w:sz="4" w:space="0" w:color="E0E0E0"/>'
        f'</w:tblBorders>'
    )
    tblPr.append(borders)

def get_char_diff(old_text, new_text):
    matcher = difflib.SequenceMatcher(None, old_text, new_text)
    opcodes = matcher.get_opcodes()
    diff_runs = []
    for tag, i1, i2, j1, j2 in opcodes:
        if tag == 'equal':
            diff_runs.append(('equal', old_text[i1:i2]))
        elif tag == 'delete':
            diff_runs.append(('delete', old_text[i1:i2]))
        elif tag == 'insert':
            diff_runs.append(('insert', new_text[j1:j2]))
        elif tag == 'replace':
            diff_runs.append(('delete', old_text[i1:i2]))
            diff_runs.append(('insert', new_text[j1:j2]))

    merged_runs = []
    for tag, text in diff_runs:
        if not text:
            continue
        if merged_runs and merged_runs[-1][0] == tag:
            merged_runs[-1] = (tag, merged_runs[-1][1] + text)
        else:
            merged_runs.append((tag, text))
    return merged_runs

def is_heading(paragraph):
    text = paragraph.text.strip()
    if not text:
        return False
    heading_patterns = [
        r'^제\s*\d+\s*[조장항장회]\b', r'^(?:\d+\.)+\d*\s*', r'^[가-힣]\.\s*',
        r'^(?:\d+|[가-힣]|[a-zA-Z])\)\s*', r'^\(\s*(?:\d+|[가-힣]|[a-zA-Z])\s*\)\s*',
        r'^[①-⑳]', r'^\[[^\]]+\]', r'^(?:■|●|▲|◆|○|▶|◈|※|•|\*)\s*'
    ]
    return any(re.match(pattern, text) for pattern in heading_patterns)

def get_table_summary(table):
    rows_text = []
    for row in table.rows:
        row_text = " | ".join(cell.text.strip() for cell in row.cells)
        rows_text.append(row_text)
    return "\n".join(rows_text)

def iter_block_items(parent):
    if isinstance(parent, docx.document.Document):
        parent_elm = parent.element.body
    elif isinstance(parent, docx.table._Cell):
        parent_elm = parent._tc
    else:
        raise ValueError("Invalid parent type")

    for child in parent_elm.iterchildren():
        if isinstance(child, docx.oxml.text.paragraph.CT_P):
            yield docx.text.paragraph.Paragraph(child, parent)
        elif isinstance(child, docx.oxml.table.CT_Tbl):
            yield docx.table.Table(child, parent)

def apply_old_inline_diff(paragraph, diff_runs):
    p_elm = paragraph._p
    for run in list(paragraph.runs):
        p_elm.remove(run._r)
    for tag, text in diff_runs:
        if tag == 'equal':
            paragraph.add_run(text)
        elif tag == 'delete':
            run = paragraph.add_run(text)
            run.font.color.rgb = COLOR_RED

def apply_new_inline_diff(paragraph, diff_runs):
    p_elm = paragraph._p
    for run in list(paragraph.runs):
        p_elm.remove(run._r)
    for tag, text in diff_runs:
        if tag == 'equal':
            paragraph.add_run(text)
        elif tag == 'insert':
            run = paragraph.add_run(text)
            run.font.color.rgb = COLOR_BLUE

def color_entire_block(block, rgb_color):
    if isinstance(block, docx.text.paragraph.Paragraph):
        for run in block.runs:
            run.font.color.rgb = rgb_color
    else:
        for row in block.rows:
            for cell in row.cells:
                for p in cell.paragraphs:
                    for run in p.runs:
                        run.font.color.rgb = rgb_color

def match_replaced_ranges(blocks1_range, blocks2_range):
    paired = []
    min_len = min(len(blocks1_range), len(blocks2_range))
    for k in range(min_len):
        paired.append((blocks1_range[k], blocks2_range[k]))
    if len(blocks1_range) > min_len:
        for k in range(min_len, len(blocks1_range)):
            paired.append((blocks1_range[k], None))
    elif len(blocks2_range) > min_len:
        for k in range(min_len, len(blocks2_range)):
            paired.append((None, blocks2_range[k]))
    return paired

def get_regulatory_location(blocks, current_idx):
    for idx in range(current_idx - 1, -1, -1):
        b = blocks[idx]
        if isinstance(b, docx.text.paragraph.Paragraph):
            t = b.text.strip()
            if t and (re.match(r'^\d+\.', t) or re.match(r'^(?:■|◆|▶|◈)\s*', t) or is_heading(b)):
                return t[:30]
    return "본문"

def copy_group_blocks_to_cell(group, dest_cell, is_old):
    tc = dest_cell._tc
    for child in list(tc):
        if child.tag.endswith('Pr'):
            continue
        tc.remove(child)

    has_any_block = False
    for item in group['items']:
        src_block = item['old_block'] if is_old else item['new_block']
        if src_block is not None:
            has_any_block = True
            if isinstance(src_block, docx.text.paragraph.Paragraph):
                block_xml = copy.deepcopy(src_block._p)
                tc.append(block_xml)
            elif isinstance(src_block, docx.table.Table):
                block_xml = copy.deepcopy(src_block._tbl)
                tc.append(block_xml)

    if not has_any_block:
        p_none = OxmlElement('w:p')
        r_none = OxmlElement('w:r')
        t_none = OxmlElement('w:t')
        t_none.text = "(없음)"
        r_none.append(t_none)
        p_none.append(r_none)
        tc.append(p_none)

def write_records_to_table(comparison_records, table):
    grouped_records = []
    for record in comparison_records:
        if grouped_records and grouped_records[-1]['loc'] == record['loc']:
            grouped_records[-1]['items'].append(record)
        else:
            grouped_records.append({'loc': record['loc'], 'items': [record]})

    for r_idx, group in enumerate(grouped_records):
        row_cells = table.add_row().cells
        row_cells[0].text = group['loc']
        set_cell_margins(row_cells[0])

        copy_group_blocks_to_cell(group, row_cells[1], is_old=True)
        set_cell_margins(row_cells[1])

        copy_group_blocks_to_cell(group, row_cells[2], is_old=False)
        set_cell_margins(row_cells[2])

        row_cells[3].text = ""
        set_cell_margins(row_cells[3])

# ----------------------------------------------------
# 문서 비교 핵심 처리 함수
# ----------------------------------------------------
def process_doc_comparison(doc1, doc2):
    blocks1 = list(iter_block_items(doc1))
    blocks2 = list(iter_block_items(doc2))

    blocks1_repr = [('p', b.text.strip()) if isinstance(b, docx.text.paragraph.Paragraph) else ('t', get_table_summary(b)) for b in blocks1]
    blocks2_repr = [('p', b.text.strip()) if isinstance(b, docx.text.paragraph.Paragraph) else ('t', get_table_summary(b)) for b in blocks2]

    matcher = difflib.SequenceMatcher(None, blocks1_repr, blocks2_repr)
    opcodes = matcher.get_opcodes()
    comparison_records = []

    for tag, i1, i2, j1, j2 in opcodes:
        if tag == 'equal':
            continue
        elif tag == 'insert':
            for j in range(j1, j2):
                color_entire_block(blocks2[j], COLOR_BLUE)
                comparison_records.append({'loc': get_regulatory_location(blocks2, j), 'old_block': None, 'new_block': blocks2[j]})
        elif tag == 'delete':
            for i in range(i1, i2):
                color_entire_block(blocks1[i], COLOR_RED)
                comparison_records.append({'loc': get_regulatory_location(blocks1, i), 'old_block': blocks1[i], 'new_block': None})
        elif tag == 'replace':
            old_subset, new_subset = blocks1[i1:i2], blocks2[j1:j2]
            paired = match_replaced_ranges(old_subset, new_subset)
            for old_b, new_b in paired:
                if old_b and isinstance(old_b, docx.text.paragraph.Paragraph) and new_b and isinstance(new_b, docx.text.paragraph.Paragraph):
                    diff_runs = get_char_diff(old_b.text.strip(), new_b.text.strip())
                    apply_old_inline_diff(old_b, diff_runs)
                    apply_new_inline_diff(new_b, diff_runs)
                else:
                    if old_b: color_entire_block(old_b, COLOR_RED)
                    if new_b: color_entire_block(new_b, COLOR_BLUE)
                loc = get_regulatory_location(blocks2, j1) if new_b else get_regulatory_location(blocks1, i1)
                comparison_records.append({'loc': loc, 'old_block': old_b, 'new_block': new_b})

    return comparison_records

# ----------------------------------------------------
# API 엔드포인트
# ----------------------------------------------------
@app.get("/ping")
def ping():
    return {"status": "ok", "message": "Server awake!"}

@app.post("/compare")
async def compare_documents(
    old_file: UploadFile = File(...),
    new_file: UploadFile = File(...),
    authorization: str = Header(None)
):
    if authorization != f"Bearer {MY_API_KEY}":
        raise HTTPException(status_code=401, detail="인증 실패")

    old_bytes = await old_file.read()
    new_bytes = await new_file.read()

    doc1 = docx.Document(io.BytesIO(old_bytes))
    doc2 = docx.Document(io.BytesIO(new_bytes))

    comparison_records = process_doc_comparison(doc1, doc2)

    # 결과 표 문서 구성
    doc_table = docx.Document()
    title_p = doc_table.add_paragraph()
    title_run = title_p.add_run("문서 변경대비표")
    title_run.font.size = Pt(14)
    title_run.font.bold = True

    table = doc_table.add_table(rows=1, cols=4)
    table.style = 'Table Grid'
    set_table_borders(table)

    hdr_row = table.rows[0]
    headers = ["위치", "구버전", "신버전", "비고"]
    for idx, text in enumerate(headers):
        cell = hdr_row.cells[idx]
        cell.text = text
        set_cell_background(cell, HEX_HEADER_BG)
        set_cell_margins(cell)

    if comparison_records:
        write_records_to_table(comparison_records, table)

    col_widths = [Inches(1.5), Inches(3.0), Inches(3.0), Inches(1.0)]
    for row in table.rows:
        for idx, width in enumerate(col_widths):
            row.cells[idx].width = width

    output_stream = io.BytesIO()
    doc_table.save(output_stream)
    output_stream.seek(0)

    return StreamingResponse(
        output_stream,
        media_type="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        headers={"Content-Disposition": "attachment; filename=comparison_result.docx"}
    )
