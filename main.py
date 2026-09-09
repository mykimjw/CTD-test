import os
import re
import io
import json
import copy
import difflib
from collections import OrderedDict
from typing import Any, Dict, List, Optional, Tuple

import docx
from docx import Document
from docx.shared import RGBColor, Pt, Inches
from docx.enum.section import WD_ORIENT
from docx.oxml import OxmlElement, parse_xml
from docx.oxml.ns import qn, nsdecls
from fastapi import FastAPI, Request, HTTPException
from fastapi.responses import StreamingResponse


# ============================================================
# FastAPI
# ============================================================

app = FastAPI(title="CTD Compare API", version="3.0")

API_KEY = os.environ.get("API_KEY", "").strip()

COLOR_RED = RGBColor(255, 0, 0)
COLOR_BLUE = RGBColor(0, 0, 255)
COLOR_BLACK = RGBColor(0, 0, 0)
COLOR_GRAY = RGBColor(128, 128, 128)

HEX_HEADER_BG = "E6EEF8"


@app.get("/ping")
def ping():
    return {"ok": True, "service": "ctd-compare-v3"}


def verify_auth(request: Request):
    if not API_KEY:
        raise HTTPException(
            status_code=500,
            detail="Render 환경변수 API_KEY가 설정되어 있지 않습니다."
        )

    auth = request.headers.get("authorization", "")
    expected = f"Bearer {API_KEY}"

    if auth != expected:
        raise HTTPException(status_code=401, detail="Unauthorized")


# ============================================================
# 기본 유틸
# ============================================================

def normalize_space(text: str) -> str:
    return re.sub(r"\s+", " ", str(text or "")).strip()


def safe_text(text: str) -> str:
    return str(text or "").replace("\x00", "")


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


def paragraph_has_drawing(p: docx.text.paragraph.Paragraph) -> bool:
    for el in p._p.iter():
        if el.tag == qn("w:drawing") or el.tag == qn("w:pict"):
            return True
    return False


def get_table_summary(table: docx.table.Table) -> str:
    rows = []
    for row in table.rows:
        rows.append(" | ".join(normalize_space(c.text) for c in row.cells))
    return "\n".join(rows)


def block_repr(block):
    if isinstance(block, docx.text.paragraph.Paragraph):
        if paragraph_has_drawing(block) and not normalize_space(block.text):
            return ("img", "[IMAGE]")
        return ("p", normalize_space(block.text))

    if isinstance(block, docx.table.Table):
        return ("t", get_table_summary(block))

    return ("x", "")


def set_run_east_asia(run, font_name="맑은 고딕"):
    run.font.name = font_name

    rPr = run._r.get_or_add_rPr()
    rFonts = rPr.rFonts

    if rFonts is None:
        rFonts = OxmlElement("w:rFonts")
        rPr.append(rFonts)

    rFonts.set(qn("w:ascii"), font_name)
    rFonts.set(qn("w:hAnsi"), font_name)
    rFonts.set(qn("w:eastAsia"), font_name)


def set_all_fonts_in_cell(cell, size=8.5):
    for p in cell.paragraphs:
        for run in p.runs:
            set_run_east_asia(run)
            run.font.size = Pt(size)

    for table in cell.tables:
        for row in table.rows:
            for c in row.cells:
                set_all_fonts_in_cell(c, size=size)


def set_cell_margins(cell, top=80, bottom=80, left=100, right=100):
    tcPr = cell._tc.get_or_add_tcPr()
    tcMar = tcPr.find(qn("w:tcMar"))

    if tcMar is None:
        tcMar = OxmlElement("w:tcMar")
        tcPr.append(tcMar)

    values = {
        "top": top,
        "bottom": bottom,
        "left": left,
        "right": right,
    }

    for k, v in values.items():
        node = tcMar.find(qn(f"w:{k}"))
        if node is None:
            node = OxmlElement(f"w:{k}")
            tcMar.append(node)

        node.set(qn("w:w"), str(v))
        node.set(qn("w:type"), "dxa")


def set_cell_background(cell, hex_color):
    tcPr = cell._tc.get_or_add_tcPr()
    shd = tcPr.find(qn("w:shd"))

    if shd is None:
        shd = parse_xml(
            f'<w:shd {nsdecls("w")} w:fill="{hex_color}"/>'
        )
        tcPr.append(shd)
    else:
        shd.set(qn("w:fill"), hex_color)


def set_table_borders(table):
    tblPr = table._tbl.tblPr

    old = tblPr.find(qn("w:tblBorders"))
    if old is not None:
        tblPr.remove(old)

    borders = parse_xml(
        f'<w:tblBorders {nsdecls("w")}>'
        f'<w:top w:val="single" w:sz="4" w:space="0" w:color="808080"/>'
        f'<w:bottom w:val="single" w:sz="4" w:space="0" w:color="808080"/>'
        f'<w:left w:val="single" w:sz="4" w:space="0" w:color="808080"/>'
        f'<w:right w:val="single" w:sz="4" w:space="0" w:color="808080"/>'
        f'<w:insideH w:val="single" w:sz="4" w:space="0" w:color="A0A0A0"/>'
        f'<w:insideV w:val="single" w:sz="4" w:space="0" w:color="A0A0A0"/>'
        f'</w:tblBorders>'
    )
    tblPr.append(borders)


# ============================================================
# Diff
# ============================================================

def get_char_diff(old_text: str, new_text: str):
    matcher = difflib.SequenceMatcher(None, old_text, new_text)

    result = []
    for tag, i1, i2, j1, j2 in matcher.get_opcodes():
        if tag == "equal":
            result.append(("equal", old_text[i1:i2]))
        elif tag == "delete":
            result.append(("delete", old_text[i1:i2]))
        elif tag == "insert":
            result.append(("insert", new_text[j1:j2]))
        elif tag == "replace":
            result.append(("delete", old_text[i1:i2]))
            result.append(("insert", new_text[j1:j2]))

    merged = []
    for tag, text in result:
        if not text:
            continue

        if merged and merged[-1][0] == tag:
            merged[-1] = (tag, merged[-1][1] + text)
        else:
            merged.append((tag, text))

    return merged


def clear_paragraph_runs(p):
    for run in list(p.runs):
        p._p.remove(run._r)


def append_diff_runs(p, diffs, side: str):
    clear_paragraph_runs(p)

    for tag, text in diffs:
        if side == "old" and tag == "insert":
            continue
        if side == "new" and tag == "delete":
            continue

        run = p.add_run(text)
        set_run_east_asia(run)

        if side == "old" and tag == "delete":
            run.font.color.rgb = COLOR_RED
        elif side == "new" and tag == "insert":
            run.font.color.rgb = COLOR_BLUE
        else:
            run.font.color.rgb = COLOR_BLACK


def apply_inline_diff_pair(old_p, new_p):
    old_text = old_p.text
    new_text = new_p.text
    diffs = get_char_diff(old_text, new_text)

    append_diff_runs(old_p, diffs, "old")
    append_diff_runs(new_p, diffs, "new")


def color_entire_block(block, color):
    if isinstance(block, docx.text.paragraph.Paragraph):
        if not block.runs and block.text:
            block.add_run(block.text)

        for run in block.runs:
            run.font.color.rgb = color
            set_run_east_asia(run)

    elif isinstance(block, docx.table.Table):
        for row in block.rows:
            for cell in row.cells:
                for p in cell.paragraphs:
                    for run in p.runs:
                        run.font.color.rgb = color
                        set_run_east_asia(run)


def match_replaced_ranges(old_range, new_range):
    paired = []
    n = min(len(old_range), len(new_range))

    for i in range(n):
        paired.append((old_range[i], new_range[i]))

    for i in range(n, len(old_range)):
        paired.append((old_range[i], None))

    for i in range(n, len(new_range)):
        paired.append((None, new_range[i]))

    return paired


def compare_and_mark_tables(t1, t2):
    old_rows = [" | ".join(normalize_space(c.text) for c in r.cells)
                for r in t1.rows]
    new_rows = [" | ".join(normalize_space(c.text) for c in r.cells)
                for r in t2.rows]

    matcher = difflib.SequenceMatcher(None, old_rows, new_rows)

    for tag, i1, i2, j1, j2 in matcher.get_opcodes():
        if tag == "equal":
            continue

        if tag == "delete":
            for r in t1.rows[i1:i2]:
                for c in r.cells:
                    color_entire_block_table_cell(c, COLOR_RED)

        elif tag == "insert":
            for r in t2.rows[j1:j2]:
                for c in r.cells:
                    color_entire_block_table_cell(c, COLOR_BLUE)

        elif tag == "replace":
            old_part = t1.rows[i1:i2]
            new_part = t2.rows[j1:j2]
            n = min(len(old_part), len(new_part))

            for k in range(n):
                r1 = old_part[k]
                r2 = new_part[k]
                c_n = min(len(r1.cells), len(r2.cells))

                for ci in range(c_n):
                    c1 = r1.cells[ci]
                    c2 = r2.cells[ci]

                    p_n = min(len(c1.paragraphs), len(c2.paragraphs))

                    for pi in range(p_n):
                        apply_inline_diff_pair(
                            c1.paragraphs[pi],
                            c2.paragraphs[pi]
                        )

                    for pi in range(p_n, len(c1.paragraphs)):
                        for run in c1.paragraphs[pi].runs:
                            run.font.color.rgb = COLOR_RED

                    for pi in range(p_n, len(c2.paragraphs)):
                        for run in c2.paragraphs[pi].runs:
                            run.font.color.rgb = COLOR_BLUE

                for ci in range(c_n, len(r1.cells)):
                    color_entire_block_table_cell(r1.cells[ci], COLOR_RED)

                for ci in range(c_n, len(r2.cells)):
                    color_entire_block_table_cell(r2.cells[ci], COLOR_BLUE)

            for r in old_part[n:]:
                for c in r.cells:
                    color_entire_block_table_cell(c, COLOR_RED)

            for r in new_part[n:]:
                for c in r.cells:
                    color_entire_block_table_cell(c, COLOR_BLUE)


def color_entire_block_table_cell(cell, color):
    for p in cell.paragraphs:
        for run in p.runs:
            run.font.color.rgb = color
            set_run_east_asia(run)


# ============================================================
# CTD 위치 파싱
# ============================================================

CTD_RE = re.compile(
    r'(?<![\w.])'
    r'(2\.3\.[A-Za-z]\.\d+(?:\.\d+)*(?:-\d+)?)'
)


def extract_ctd_code(text: str) -> Optional[str]:
    m = CTD_RE.search(normalize_space(text))
    return m.group(1) if m else None


def ctd_depth(code: str) -> int:
    # 2.3.S.2.1.1 -> ["2","3","S","2","1","1"] => 6
    return len(re.split(r"[.-]", code))


def is_ctd_heading_text(text: str) -> bool:
    text = normalize_space(text)
    if not text:
        return False

    code = extract_ctd_code(text)
    if not code:
        return False

    # 문장 중간의 참조번호가 아니라 제목처럼 앞쪽에 있는 경우
    return text.startswith(code) or text.startswith(code + ".")


def find_preceding_table_title(blocks, current_idx):
    patterns = [
        r"^\s*표\s*2\.3\.",
        r"^\s*Table\s*2\.3\.",
        r"^\s*그림\s*2\.3\.",
        r"^\s*Figure\s*2\.3\.",
    ]

    for idx in range(current_idx - 1, -1, -1):
        b = blocks[idx]

        if not isinstance(b, docx.text.paragraph.Paragraph):
            continue

        text = normalize_space(b.text)
        if not text:
            continue

        if any(re.match(pat, text, re.I) for pat in patterns):
            return text

        if is_ctd_heading_text(text):
            break

    return None


def get_regulatory_location(blocks, current_idx):
    current = blocks[current_idx]
    path_by_depth = {}

    # 현재 블록 자체가 CTD 제목이면 포함
    if isinstance(current, docx.text.paragraph.Paragraph):
        t = normalize_space(current.text)
        code = extract_ctd_code(t)

        if code and is_ctd_heading_text(t):
            path_by_depth[ctd_depth(code)] = t

    table_title = None
    if isinstance(current, docx.table.Table):
        table_title = find_preceding_table_title(blocks, current_idx)

    # 역방향으로 상위 CTD 제목 추적
    for idx in range(current_idx - 1, -1, -1):
        b = blocks[idx]

        if not isinstance(b, docx.text.paragraph.Paragraph):
            continue

        text = normalize_space(b.text)
        if not text:
            continue

        code = extract_ctd_code(text)
        if not code or not is_ctd_heading_text(text):
            continue

        depth = ctd_depth(code)

        if depth not in path_by_depth:
            path_by_depth[depth] = text

        # 2.3.P.2 / 2.3.S.2 등 최상위 소그룹을 찾으면 충분
        if depth <= 4:
            break

    path = [
        path_by_depth[d]
        for d in sorted(path_by_depth.keys())
    ]

    if table_title:
        path.append(table_title)

    return " > ".join(path) if path else "문서 시작"


def top_ctd_group(location: str, doc_subtype: str) -> str:
    source = location + " " + doc_subtype
    code = extract_ctd_code(source)

    if not code:
        return normalize_space(doc_subtype) or "기타"

    parts = code.split(".")
    if len(parts) >= 4:
        return ".".join(parts[:4])

    return code


def natural_ctd_sort_key(code: str):
    m = re.match(r"^(\d+)\.(\d+)\.([A-Za-z]+)\.(\d+)", code)

    if not m:
        return (9, 999, 999, "Z", 999, code)

    return (
        0,
        int(m.group(1)),
        int(m.group(2)),
        m.group(3).upper(),
        int(m.group(4)),
        code
    )


# ============================================================
# 비교 레코드 생성
# ============================================================

def compare_documents(old_bytes: Optional[bytes],
                      new_bytes: bytes,
                      doc_subtype: str,
                      is_new_doc: bool):

    new_doc = Document(io.BytesIO(new_bytes))
    new_blocks = list(iter_block_items(new_doc))

    if not old_bytes or is_new_doc:
        records = []

        # 신규 문서 전체를 다 쓰지 않고 CTD 제목/실내용 중심으로 등록
        for j, b in enumerate(new_blocks):
            rep = block_repr(b)

            # 빈 문단 제외
            if rep[0] == "p" and not rep[1]:
                continue

            color_entire_block(b, COLOR_BLUE)

            loc = get_regulatory_location(new_blocks, j)

            records.append({
                "loc": loc,
                "old_block": None,
                "new_block": b,
                "type": "insert",
                "doc_subtype": doc_subtype,
                "old_index": None,
                "new_index": j,
            })

        return records

    old_doc = Document(io.BytesIO(old_bytes))
    old_blocks = list(iter_block_items(old_doc))

    old_repr = [block_repr(b) for b in old_blocks]
    new_repr = [block_repr(b) for b in new_blocks]

    matcher = difflib.SequenceMatcher(None, old_repr, new_repr)
    opcodes = matcher.get_opcodes()

    records = []

    # 먼저 색상 적용 + 레코드 생성
    for tag, i1, i2, j1, j2 in opcodes:
        if tag == "equal":
            continue

        if tag == "insert":
            for j in range(j1, j2):
                b = new_blocks[j]
                color_entire_block(b, COLOR_BLUE)

                records.append({
                    "loc": get_regulatory_location(new_blocks, j),
                    "old_block": None,
                    "new_block": b,
                    "type": "insert",
                    "doc_subtype": doc_subtype,
                    "old_index": None,
                    "new_index": j,
                })

        elif tag == "delete":
            for i in range(i1, i2):
                b = old_blocks[i]
                color_entire_block(b, COLOR_RED)

                records.append({
                    "loc": get_regulatory_location(old_blocks, i),
                    "old_block": b,
                    "new_block": None,
                    "type": "delete",
                    "doc_subtype": doc_subtype,
                    "old_index": i,
                    "new_index": None,
                })

        elif tag == "replace":
            old_part = old_blocks[i1:i2]
            new_part = new_blocks[j1:j2]
            pairs = match_replaced_ranges(old_part, new_part)

            for k, (ob, nb) in enumerate(pairs):
                oi = i1 + k if ob is not None and k < len(old_part) else None
                nj = j1 + k if nb is not None and k < len(new_part) else None

                if ob is not None and nb is not None:
                    if isinstance(ob, docx.text.paragraph.Paragraph) and \
                       isinstance(nb, docx.text.paragraph.Paragraph):
                        apply_inline_diff_pair(ob, nb)

                    elif isinstance(ob, docx.table.Table) and \
                         isinstance(nb, docx.table.Table):
                        compare_and_mark_tables(ob, nb)

                    else:
                        color_entire_block(ob, COLOR_RED)
                        color_entire_block(nb, COLOR_BLUE)

                elif ob is not None:
                    color_entire_block(ob, COLOR_RED)

                elif nb is not None:
                    color_entire_block(nb, COLOR_BLUE)

                if nb is not None and nj is not None:
                    loc = get_regulatory_location(new_blocks, nj)
                elif ob is not None and oi is not None:
                    loc = get_regulatory_location(old_blocks, oi)
                else:
                    loc = doc_subtype

                records.append({
                    "loc": loc,
                    "old_block": ob,
                    "new_block": nb,
                    "type": "replace",
                    "doc_subtype": doc_subtype,
                    "old_index": oi,
                    "new_index": nj,
                })

    return records


# ============================================================
# 표/문단 축약
# ============================================================

def is_run_modified(run_xml):
    rPr = run_xml.find(qn("w:rPr"))
    if rPr is None:
        return False

    color = rPr.find(qn("w:color"))
    if color is None:
        return False

    value = color.get(qn("w:val"))
    return bool(value and value.upper() in ("FF0000", "0000FF"))


def is_row_modified(tr_xml):
    for r in tr_xml.iter(qn("w:r")):
        if is_run_modified(r):
            return True
    return False


def make_ellipsis_paragraph(text="(중략)"):
    p = OxmlElement("w:p")
    r = OxmlElement("w:r")
    rPr = OxmlElement("w:rPr")

    color = OxmlElement("w:color")
    color.set(qn("w:val"), "808080")
    rPr.append(color)

    r.append(rPr)

    t = OxmlElement("w:t")
    t.text = text
    r.append(t)

    p.append(r)
    return p


def create_ellipsis_row(template_tr, text="(표 생략)"):
    tr = copy.deepcopy(template_tr)

    cells = list(tr.iter(qn("w:tc")))

    for idx, tc in enumerate(cells):
        # 기존 문단/표 제거
        for child in list(tc):
            if child.tag == qn("w:tcPr"):
                continue
            tc.remove(child)

        p = OxmlElement("w:p")
        r = OxmlElement("w:r")
        rPr = OxmlElement("w:rPr")

        color = OxmlElement("w:color")
        color.set(qn("w:val"), "808080")
        rPr.append(color)
        r.append(rPr)

        t = OxmlElement("w:t")
        t.text = text if idx == 0 else ""
        r.append(t)

        p.append(r)
        tc.append(p)

    return tr


def prune_table_xml(tbl_xml):
    rows = tbl_xml.findall(qn("w:tr"))

    if len(rows) <= 4:
        return

    modified = []
    for idx, row in enumerate(rows):
        if idx == 0:
            continue
        if is_row_modified(row):
            modified.append(idx)

    if not modified:
        # 실제 변경 색상이 없으면 너무 큰 표는 생략 표시
        for row in rows[1:]:
            tbl_xml.remove(row)
        tbl_xml.append(create_ellipsis_row(rows[0]))
        return

    keep = {0}
    for idx in modified:
        keep.add(idx)
        if idx - 1 > 0:
            keep.add(idx - 1)
        if idx + 1 < len(rows):
            keep.add(idx + 1)

    new_rows = []
    last = -1

    for idx in sorted(keep):
        if last >= 0 and idx - last > 1:
            new_rows.append(create_ellipsis_row(rows[idx]))
        new_rows.append(rows[idx])
        last = idx

    if last < len(rows) - 1:
        new_rows.append(create_ellipsis_row(rows[-1]))

    for row in rows:
        tbl_xml.remove(row)

    for row in new_rows:
        tbl_xml.append(row)


def fit_table_to_width(tbl_xml, target_width_dxa):
    tblGrid = tbl_xml.find(qn("w:tblGrid"))
    widths = []

    if tblGrid is not None:
        for col in tblGrid.findall(qn("w:gridCol")):
            v = col.get(qn("w:w"))
            widths.append(float(v) if v else 1000.0)

    if not widths:
        first_row = tbl_xml.find(qn("w:tr"))
        if first_row is not None:
            count = len(first_row.findall(qn("w:tc")))
            if count:
                widths = [1000.0] * count

    if not widths:
        return

    total = sum(widths)
    if total <= 0:
        return

    scale = target_width_dxa / total
    scaled = [max(120, int(x * scale)) for x in widths]

    if tblGrid is not None:
        for col in list(tblGrid.findall(qn("w:gridCol"))):
            tblGrid.remove(col)

        for w in scaled:
            col = OxmlElement("w:gridCol")
            col.set(qn("w:w"), str(w))
            tblGrid.append(col)

    tblPr = tbl_xml.find(qn("w:tblPr"))
    if tblPr is not None:
        tblW = tblPr.find(qn("w:tblW"))

        if tblW is None:
            tblW = OxmlElement("w:tblW")
            tblPr.append(tblW)

        tblW.set(qn("w:type"), "dxa")
        tblW.set(qn("w:w"), str(target_width_dxa))


# ============================================================
# 출력 셀 작성
# ============================================================

def clear_cell(cell):
    tc = cell._tc

    for child in list(tc):
        if child.tag == qn("w:tcPr"):
            continue
        tc.remove(child)


def append_text_paragraph(cell, text, bold=False, color=None,
                          size=8.5, center=False):
    p = cell.add_paragraph()

    if center:
        p.alignment = 1

    run = p.add_run(text)
    set_run_east_asia(run)
    run.font.size = Pt(size)
    run.font.bold = bold

    if color:
        run.font.color.rgb = color

    return p


def location_parts_for_cell(location, top_group):
    parts = [normalize_space(x) for x in location.split(" > ") if normalize_space(x)]

    result = []
    for p in parts:
        code = extract_ctd_code(p)

        if code and top_group and code.startswith(top_group):
            # 최상위 top_group 자체는 1열에 있으므로 제외
            if code == top_group:
                continue

        result.append(p)

    return result


def same_heading_prefix(a: List[str], b: List[str]) -> int:
    n = min(len(a), len(b))

    for i in range(n):
        if a[i] != b[i]:
            return i

    return n


def block_to_cell(cell, block, side):
    if block is None:
        marker = "(신규)" if side == "old" else "(삭제)"
        append_text_paragraph(
            cell,
            marker,
            bold=True,
            color=COLOR_RED if side == "old" else COLOR_BLUE,
            size=8.5,
            center=True
        )
        return

    if isinstance(block, docx.text.paragraph.Paragraph):
        if paragraph_has_drawing(block):
            # 그림은 용량/레이아웃 문제 방지를 위해 정답 예시처럼 생략 표기
            append_text_paragraph(
                cell,
                "(그림 생략)",
                color=COLOR_GRAY,
                size=8.5,
                center=True
            )

            # 그림 제목 텍스트가 있으면 같이 보존
            if normalize_space(block.text):
                xml = copy.deepcopy(block._p)
                cell._tc.append(xml)

            return

        if not normalize_space(block.text):
            return

        xml = copy.deepcopy(block._p)
        cell._tc.append(xml)
        return

    if isinstance(block, docx.table.Table):
        xml = copy.deepcopy(block._tbl)
        prune_table_xml(xml)

        target_width = 4400
        if cell.width:
            try:
                target_width = max(2000, int(cell.width.inches * 1440) - 150)
            except Exception:
                pass

        fit_table_to_width(xml, target_width)
        cell._tc.append(xml)


def write_side_records(cell, records, side, top_group):
    clear_cell(cell)

    last_headings = []
    last_loc = None

    for idx, record in enumerate(records):
        loc = record["loc"]
        headings = location_parts_for_cell(loc, top_group)

        # 위치가 크게 바뀌면 중략
        if idx > 0 and loc != last_loc:
            append_text_paragraph(
                cell,
                "(중략)",
                color=COLOR_GRAY,
                size=8.0,
                center=True
            )

        common = same_heading_prefix(last_headings, headings)

        for level in range(common, len(headings)):
            text = headings[level]

            prefix = "■ " if level == 0 else ("  " * level + "└ ")
            append_text_paragraph(
                cell,
                prefix + text,
                bold=True,
                color=COLOR_BLACK,
                size=8.5
            )

        block = record["old_block"] if side == "old" else record["new_block"]
        block_to_cell(cell, block, side)

        if headings:
            last_headings = headings

        last_loc = loc

    # cell은 반드시 마지막 문단이 있어야 Word에서 안정적
    cell._tc.append(OxmlElement("w:p"))
    set_all_fonts_in_cell(cell, 8.5)
    set_cell_margins(cell)


# ============================================================
# 변경 사유 자동 추론
# ============================================================

def record_text(record):
    parts = []

    for block in (record.get("old_block"), record.get("new_block")):
        if isinstance(block, docx.text.paragraph.Paragraph):
            parts.append(block.text)
        elif isinstance(block, docx.table.Table):
            parts.append(get_table_summary(block))

    parts.append(record.get("loc", ""))
    parts.append(record.get("doc_subtype", ""))

    return normalize_space(" ".join(parts))


def infer_reason(records):
    text = " ".join(record_text(r) for r in records)

    rules = [
        (
            ["제조처", "제조원", "동방에프티엘"],
            "주성분 제조원 추가 및 관련 자료 반영"
        ),
        (
            ["공정 밸리데이션", "PV", "PVP"],
            "주성분 제조원 추가에 따른 공정 밸리데이션 자료 반영"
        ),
        (
            ["안정성", "장기", "가속"],
            "주성분 제조원 추가로 인한 안정성 시험 자료 추가"
        ),
        (
            ["주소", "전화번호", "팩스번호"],
            "제조원 추가에 따른 제조원 정보 기재"
        ),
        (
            ["원료관리", "Control of Materials"],
            "주성분 제조원 추가에 따른 원료관리 자료 반영"
        ),
        (
            ["제조공정", "Manufacturing Process"],
            "주성분 제조원 추가에 따른 제조공정 자료 반영"
        ),
        (
            ["시험결과", "시험 결과", "Batch no", "Batch No"],
            "변경 사항에 따른 시험 결과 반영"
        ),
        (
            ["NMR", "XRD", "IR", "Mass", "Elemental"],
            "주성분 제조원 추가에 따른 원료 물리화학적 특성 자료 반영"
        ),
    ]

    matched = []

    for keywords, reason in rules:
        if any(k.lower() in text.lower() for k in keywords):
            if reason not in matched:
                matched.append(reason)

    if not matched:
        types = {r.get("type") for r in records}

        if types == {"insert"}:
            return "신규 자료 추가"
        if types == {"delete"}:
            return "기존 자료 삭제"

        return "변경사항 반영"

    # 너무 길어지는 것을 막기 위해 최대 2개
    return "\n".join(matched[:2])


# ============================================================
# 템플릿
# ============================================================

def replace_placeholder_in_paragraph(p, placeholder, value):
    if placeholder not in p.text:
        return

    # placeholder가 한 run에 있으면 스타일 유지
    for run in p.runs:
        if placeholder in run.text:
            run.text = run.text.replace(placeholder, value)
            return

    # 여러 run에 찢겨있으면 첫 run에 통합
    full = p.text.replace(placeholder, value)

    if p.runs:
        p.runs[0].text = full
        for run in p.runs[1:]:
            run.text = ""
    else:
        p.add_run(full)


def replace_placeholder_in_element(container, placeholder, value):
    if hasattr(container, "paragraphs"):
        for p in container.paragraphs:
            replace_placeholder_in_paragraph(p, placeholder, value)

    if hasattr(container, "tables"):
        for table in container.tables:
            for row in table.rows:
                for cell in row.cells:
                    replace_placeholder_in_element(cell, placeholder, value)


def set_title_if_needed(doc, product_name):
    title_text = f"CTD 변경대비표[{product_name}]"

    replaced = False

    for p in doc.paragraphs:
        t = normalize_space(p.text)

        if "%제품명%" in p.text:
            replace_placeholder_in_paragraph(p, "%제품명%", product_name)
            replaced = True
            continue

        if "CTD 변경대비표" in t and "[" in t:
            # 기존 제목 전체 교체
            if p.runs:
                p.runs[0].text = title_text
                for run in p.runs[1:]:
                    run.text = ""
            else:
                p.add_run(title_text)

            replaced = True
            break

    if not replaced:
        p = doc.paragraphs[0] if doc.paragraphs else doc.add_paragraph()
        p.text = title_text

        if p.runs:
            p.runs[0].font.bold = True
            p.runs[0].font.size = Pt(12)
            set_run_east_asia(p.runs[0])


def find_or_create_result_table(doc):
    for table in doc.tables:
        if table.rows and len(table.rows[0].cells) >= 4:
            # 헤더가 구분/변경전/변경후/사유면 최우선
            header = " ".join(normalize_space(c.text) for c in table.rows[0].cells[:4])

            if all(x in header for x in ["구분", "변경", "사유"]):
                return table

    # 없으면 새로 생성
    table = doc.add_table(rows=1, cols=4)
    table.style = "Table Grid"

    headers = ["구분", "변경 전", "변경 후", "사유"]

    for idx, text in enumerate(headers):
        cell = table.rows[0].cells[idx]
        cell.text = text
        set_cell_background(cell, HEX_HEADER_BG)

        for run in cell.paragraphs[0].runs:
            run.font.bold = True
            run.font.size = Pt(9)
            set_run_east_asia(run)

    return table


def remove_all_data_rows(table):
    while len(table.rows) > 1:
        row = table.rows[-1]
        table._tbl.remove(row._tr)


def configure_page(doc):
    for section in doc.sections:
        # 템플릿이 세로라면 가로로 변경
        if section.orientation != WD_ORIENT.LANDSCAPE:
            section.orientation = WD_ORIENT.LANDSCAPE
            section.page_width, section.page_height = \
                section.page_height, section.page_width

        section.top_margin = Inches(0.35)
        section.bottom_margin = Inches(0.35)
        section.left_margin = Inches(0.35)
        section.right_margin = Inches(0.35)


def set_table_column_widths(table):
    # A4 Landscape 기준
    widths = [
        Inches(0.85),
        Inches(4.35),
        Inches(4.35),
        Inches(1.25),
    ]

    for row in table.rows:
        for idx, w in enumerate(widths):
            if idx < len(row.cells):
                row.cells[idx].width = w
                set_cell_margins(row.cells[idx])


# ============================================================
# 결과 문서 생성
# ============================================================

def build_output_doc(template_bytes: Optional[bytes],
                     product_name: str,
                     all_records: List[Dict[str, Any]]) -> bytes:

    if template_bytes:
        try:
            out_doc = Document(io.BytesIO(template_bytes))
        except Exception:
            out_doc = Document()
    else:
        out_doc = Document()

    configure_page(out_doc)

    # placeholder / header placeholder 치환
    replace_placeholder_in_element(out_doc, "%제품명%", product_name)

    for section in out_doc.sections:
        replace_placeholder_in_element(section.header, "%제품명%", product_name)

    set_title_if_needed(out_doc, product_name)

    table = find_or_create_result_table(out_doc)
    remove_all_data_rows(table)
    set_table_borders(table)

    grouped = OrderedDict()

    for record in all_records:
        top = top_ctd_group(
            record.get("loc", ""),
            record.get("doc_subtype", "")
        )
        grouped.setdefault(top, []).append(record)

    sorted_groups = sorted(
        grouped.items(),
        key=lambda x: natural_ctd_sort_key(x[0])
    )

    for top_group, records in sorted_groups:
        row = table.add_row()
        cells = row.cells

        # 완전 신규 그룹인지 판단
        is_new_group = all(r.get("old_block") is None for r in records)
        label = top_group + (" (신규)" if is_new_group else "")

        cells[0].text = label
        for run in cells[0].paragraphs[0].runs:
            run.font.bold = True
            run.font.size = Pt(8.5)
            set_run_east_asia(run)

        write_side_records(cells[1], records, "old", top_group)
        write_side_records(cells[2], records, "new", top_group)

        reason = infer_reason(records)
        cells[3].text = reason

        for p in cells[3].paragraphs:
            for run in p.runs:
                run.font.size = Pt(8.0)
                set_run_east_asia(run)

        set_cell_margins(cells[0])
        set_cell_margins(cells[3])

    set_table_column_widths(table)

    bio = io.BytesIO()
    out_doc.save(bio)
    bio.seek(0)

    return bio.getvalue()


# ============================================================
# API endpoint
# ============================================================

@app.post("/compare-product")
async def compare_product(request: Request):
    verify_auth(request)

    try:
        form = await request.form()
    except Exception as e:
        raise HTTPException(
            status_code=400,
            detail=f"multipart/form-data 파싱 실패: {e}"
        )

    product_name = normalize_space(form.get("product_name", ""))
    if not product_name:
        product_name = "제품명 미확인"

    meta_raw = form.get("meta_json")
    if not meta_raw:
        raise HTTPException(status_code=400, detail="meta_json이 없습니다.")

    try:
        meta = json.loads(str(meta_raw))
    except Exception as e:
        raise HTTPException(
            status_code=400,
            detail=f"meta_json 파싱 실패: {e}"
        )

    template_bytes = None
    template_obj = form.get("template_file")

    if template_obj is not None and hasattr(template_obj, "read"):
        try:
            template_bytes = await template_obj.read()
        except Exception:
            template_bytes = None

    all_records = []

    for item in meta:
        idx = int(item["index"])
        doc_subtype = str(item.get("doc_subtype", ""))
        is_new = bool(item.get("is_new", False))

        new_obj = form.get(f"new_{idx}")

        if new_obj is None or not hasattr(new_obj, "read"):
            raise HTTPException(
                status_code=400,
                detail=f"new_{idx} 파일이 없습니다."
            )

        new_bytes = await new_obj.read()

        old_bytes = None
        old_obj = form.get(f"old_{idx}")

        if old_obj is not None and hasattr(old_obj, "read"):
            old_bytes = await old_obj.read()

        try:
            records = compare_documents(
                old_bytes=old_bytes,
                new_bytes=new_bytes,
                doc_subtype=doc_subtype,
                is_new_doc=is_new or old_bytes is None
            )
        except Exception as e:
            raise HTTPException(
                status_code=500,
                detail=(
                    f"[{item.get('new_file_name', idx)}] 비교 실패: "
                    f"{type(e).__name__}: {e}"
                )
            )

        all_records.extend(records)

    if not all_records:
        raise HTTPException(
            status_code=422,
            detail="변경사항이 감지되지 않았습니다."
        )

    try:
        result_bytes = build_output_doc(
            template_bytes=template_bytes,
            product_name=product_name,
            all_records=all_records
        )
    except Exception as e:
        raise HTTPException(
            status_code=500,
            detail=f"결과 문서 생성 실패: {type(e).__name__}: {e}"
        )

    stream = io.BytesIO(result_bytes)

    return StreamingResponse(
        stream,
        media_type=(
            "application/vnd.openxmlformats-officedocument."
            "wordprocessingml.document"
        ),
        headers={
            "Content-Disposition": 'attachment; filename="ctd_compare.docx"'
        }
    )
