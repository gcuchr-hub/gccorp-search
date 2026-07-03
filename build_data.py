# ===================================================
# 과천도시공사 인사·ERP 통합검색 프로그램
# build_data.py : 업로드 폴더의 문서들을 읽어서
#                 정적 웹페이지(static_site)용 data.js로 "구워내는" 변환 스크립트
#
# 사용법:
#   1. uploads 폴더에 PDF / XLSX / HWPX / TXT 파일을 넣어둔다
#   2. 터미널에서 python build_data.py 실행
#   3. static_site/data.js (그리고 필요하면 static_site/images 안의 PDF 페이지 이미지들)가
#      새로 만들어지거나 갱신된다
#   4. static_site/index.html 을 더블클릭해서 결과 확인
#
# 문서를 추가하거나 수정했을 때마다, 이 스크립트를 다시 실행해주면 됩니다.
# ===================================================

import os
import re
import json
import zipfile
import xml.etree.ElementTree as ET

import pdfplumber
import fitz
import openpyxl

# -----------------------------------------------------
# 경로 설정
# -----------------------------------------------------
UPLOAD_FOLDER = "uploads"
OUTPUT_FOLDER = os.path.join("static_site")
OUTPUT_IMAGE_FOLDER = os.path.join(OUTPUT_FOLDER, "images")
OUTPUT_DATA_FILE = os.path.join(OUTPUT_FOLDER, "data.js")

os.makedirs(OUTPUT_FOLDER, exist_ok=True)
os.makedirs(OUTPUT_IMAGE_FOLDER, exist_ok=True)

ALLOWED_EXTENSIONS = {"pdf", "xlsx", "txt", "hwpx"}


# -----------------------------------------------------
# 분류(조례/규정/세칙/지침/기준) 관련 함수
# (app.py에서 쓰던 것과 동일한 로직)
# -----------------------------------------------------
CATEGORY_ORDER = ["조례", "규정", "세칙", "지침", "기준"]

DOC_ORDER = [
    "설립 및 운영",
    "주차장",
    "도시공원",
    "시민회관",
    "청소년",
]


def get_category_name(filename):
    # 파일명에 분류 단어가 여러 번 들어가는 경우(예: "감사규정 시행세칙")
    # 가장 뒤쪽에 나오는 단어를 진짜 분류로 인정한다
    found_index = -1
    found_keyword = None
    for keyword in CATEGORY_ORDER:
        idx = filename.rfind(keyword)
        if idx > found_index:
            found_index = idx
            found_keyword = keyword
    return found_keyword if found_keyword else "기타"


def get_category_priority(filename):
    name = get_category_name(filename)
    if name in CATEGORY_ORDER:
        return CATEGORY_ORDER.index(name)
    return len(CATEGORY_ORDER)


def get_doc_priority(filename):
    for i, keyword in enumerate(DOC_ORDER):
        if keyword in filename:
            return i
    return len(DOC_ORDER)


def no_ext(filename):
    return os.path.splitext(filename)[0]


# -----------------------------------------------------
# TXT 처리
# -----------------------------------------------------
def build_txt_doc(filepath, filename):
    with open(filepath, "r", encoding="utf-8", errors="ignore") as f:
        lines = [line.rstrip("\n") for line in f.readlines()]
    return {"type": "txt", "lines": lines}


# -----------------------------------------------------
# PDF 처리 (페이지 텍스트 + 페이지 이미지)
# -----------------------------------------------------
def render_pdf_page_image(pdf_doc, pdf_path, page_number):
    base_name = os.path.splitext(os.path.basename(pdf_path))[0]
    image_filename = f"{base_name}_p{page_number}.png"
    image_path = os.path.join(OUTPUT_IMAGE_FOLDER, image_filename)
    if not os.path.exists(image_path):
        page = pdf_doc[page_number - 1]
        zoom = 150 / 72
        matrix = fitz.Matrix(zoom, zoom)
        pix = page.get_pixmap(matrix=matrix)
        pix.save(image_path)
    return f"images/{image_filename}"


def build_pdf_doc(filepath, filename):
    pages = []
    pdf_doc = fitz.open(filepath)
    with pdfplumber.open(filepath) as pdf:
        for page_number, page in enumerate(pdf.pages, start=1):
            text = page.extract_text() or ""
            image_path = render_pdf_page_image(pdf_doc, filepath, page_number)
            pages.append({
                "pageNumber": page_number,
                "text": text,
                "image": image_path
            })
    pdf_doc.close()
    return {"type": "pdf", "pages": pages}


# -----------------------------------------------------
# XLSX 처리 (시트별 머리글 + 행 데이터)
# -----------------------------------------------------
def build_xlsx_doc(filepath, filename):
    sheets = []
    workbook = openpyxl.load_workbook(filepath, data_only=True)
    for sheet in workbook.worksheets:
        rows_iter = sheet.iter_rows(values_only=True)
        try:
            header_raw = list(next(rows_iter))
        except StopIteration:
            header_raw = []
        header = [str(v).strip() if v is not None else "" for v in header_raw]
        data_rows = []
        for row in rows_iter:
            row_values = list(row)
            if all(v is None for v in row_values):
                continue
            # JSON으로 옮길 수 있도록 값들을 문자열/숫자 형태로 정리
            clean_row = []
            for v in row_values:
                if v is None:
                    clean_row.append(None)
                else:
                    clean_row.append(v if isinstance(v, (int, float)) else str(v))
            data_rows.append(clean_row)
        sheets.append({
            "title": sheet.title,
            "header": header,
            "rows": data_rows
        })
    return {"type": "xlsx", "sheets": sheets}


# -----------------------------------------------------
# HWPX 처리 (app.py의 파싱 로직을 그대로 재사용)
# -----------------------------------------------------
def hwpx_tag(elem):
    return elem.tag.split("}")[-1]


def hwpx_escape(text):
    return text.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")


def parse_hwpx_charprops(filepath):
    charprops = {}
    with zipfile.ZipFile(filepath, "r") as zf:
        if "Contents/header.xml" not in zf.namelist():
            return charprops
        root = ET.fromstring(zf.read("Contents/header.xml"))
    for elem in root.iter():
        if hwpx_tag(elem) == "charPr":
            cid = elem.get("id")
            color = elem.get("textColor") or "#000000"
            is_bold = any(hwpx_tag(c) == "bold" for c in elem)
            charprops[cid] = {"bold": is_bold, "color": color}
    return charprops


def hwpx_has_table(run_elem):
    return any(hwpx_tag(c) == "tbl" for c in run_elem)


def hwpx_get_table(run_elem):
    for c in run_elem:
        if hwpx_tag(c) == "tbl":
            return c
    return None


def get_hwpx_run_props(run_elem, charprops):
    charpr_id = run_elem.get("charPrIDRef")
    props = charprops.get(charpr_id, {})
    return {
        "bold": bool(props.get("bold")),
        "color": props.get("color", "#000000"),
    }


def get_hwpx_t_text(t_elem):
    # <hp:t> 안에 줄바꿈(<hp:lineBreak/>) 같은 표시가 끼어있으면,
    # 그 표시 "다음"에 오는 글자는 t_elem.text가 아니라 그 표시 태그의 .tail에 들어있다.
    # (Shift+Enter로 나눈 ①②③처럼, 한 문단 안에서 줄만 바뀐 경우 이 부분을 놓치면
    #  줄바꿈 이후의 모든 글자가 통째로 사라지는 문제가 생긴다)
    parts = []
    if t_elem.text:
        parts.append(t_elem.text)
    for child in t_elem:
        parts.append("\n")
        if child.tail:
            parts.append(child.tail)
    return "".join(parts)


def get_hwpx_paragraph_runs(p_elem, charprops):
    runs = []
    for run in p_elem:
        if hwpx_tag(run) != "run":
            continue
        if hwpx_has_table(run):
            continue
        text_parts = []
        for sub in run:
            if hwpx_tag(sub) == "t":
                text = get_hwpx_t_text(sub)
                if text:
                    text_parts.append(text)
        raw_text = "".join(text_parts)
        if raw_text:
            runs.append((raw_text, get_hwpx_run_props(run, charprops)))
    return runs


def render_hwpx_runs(runs):
    plain_parts = []
    html_parts = []
    for raw_text, props in runs:
        if not raw_text:
            continue
        plain_parts.append(raw_text)
        html = hwpx_escape(raw_text).replace("\n", "<br>")
        if props.get("bold"):
            html = f"<b>{html}</b>"
        color = props.get("color", "#000000")
        if color and color.upper() not in ("#000000", "NONE"):
            html = f'<span style="color:{color}">{html}</span>'
        html_parts.append(html)
    return "".join(plain_parts).strip(), "".join(html_parts).strip()


def render_hwpx_paragraph(p_elem, charprops):
    runs = get_hwpx_paragraph_runs(p_elem, charprops)
    return render_hwpx_runs(runs)


def render_hwpx_table(tbl_elem, charprops):
    row_cnt = int(tbl_elem.get("rowCnt", "0"))
    col_cnt = int(tbl_elem.get("colCnt", "0"))
    grid = [[None for _ in range(col_cnt)] for _ in range(row_cnt)]
    plain_text_parts = []

    for tr in tbl_elem:
        if hwpx_tag(tr) != "tr":
            continue
        for tc in tr:
            if hwpx_tag(tc) != "tc":
                continue
            row_addr = col_addr = 0
            row_span = col_span = 1
            cell_html_parts = []
            for child in tc:
                ctag = hwpx_tag(child)
                if ctag == "cellAddr":
                    row_addr = int(child.get("rowAddr", "0"))
                    col_addr = int(child.get("colAddr", "0"))
                elif ctag == "cellSpan":
                    col_span = int(child.get("colSpan", "1"))
                    row_span = int(child.get("rowSpan", "1"))
                elif ctag == "subList":
                    for p in child:
                        if hwpx_tag(p) != "p":
                            continue
                        plain, html = render_hwpx_paragraph(p, charprops)
                        if html:
                            cell_html_parts.append(html)
                        if plain:
                            plain_text_parts.append(plain)
            cell_html = "<br>".join(cell_html_parts)
            if row_addr < row_cnt and col_addr < col_cnt:
                grid[row_addr][col_addr] = {
                    "html": cell_html, "rowspan": row_span, "colspan": col_span
                }

    covered = [[False] * col_cnt for _ in range(row_cnt)]
    rows_html = []
    for r in range(row_cnt):
        cells_html = []
        for c in range(col_cnt):
            if covered[r][c]:
                continue
            cell = grid[r][c]
            if cell is None:
                cells_html.append("<td></td>")
                continue
            rowspan = cell["rowspan"]
            colspan = cell["colspan"]
            for rr in range(r, min(r + rowspan, row_cnt)):
                for cc in range(c, min(c + colspan, col_cnt)):
                    covered[rr][cc] = True
            attrs = ""
            if rowspan > 1:
                attrs += f' rowspan="{rowspan}"'
            if colspan > 1:
                attrs += f' colspan="{colspan}"'
            cells_html.append(f"<td{attrs}>{cell['html']}</td>")
        rows_html.append("<tr>" + "".join(cells_html) + "</tr>")

    table_html = '<table class="hwpx-table">' + "".join(rows_html) + "</table>"
    plain_text = " ".join(plain_text_parts)
    return plain_text, table_html


def extract_hwpx_blocks(filepath):
    blocks = []
    charprops = parse_hwpx_charprops(filepath)

    with zipfile.ZipFile(filepath, "r") as zf:
        section_files = sorted(
            name for name in zf.namelist()
            if name.startswith("Contents/section") and name.endswith(".xml")
        )
        for section_name in section_files:
            root = ET.fromstring(zf.read(section_name))
            for p_elem in root:
                if hwpx_tag(p_elem) != "p":
                    continue

                table_elem = None
                for run in p_elem:
                    if hwpx_tag(run) == "run" and hwpx_has_table(run):
                        table_elem = hwpx_get_table(run)
                        break

                if table_elem is not None:
                    plain, html = render_hwpx_table(table_elem, charprops)
                    if plain.strip():
                        blocks.append({"type": "table", "plain": plain, "html": html})
                    continue

                runs = get_hwpx_paragraph_runs(p_elem, charprops)
                plain, html = render_hwpx_runs(runs)
                if plain.strip():
                    blocks.append({"type": "text", "plain": plain, "html": html, "runs": runs})

    return blocks


def render_hwpx_blocks_html(blocks):
    parts = []
    for block in blocks:
        if block["type"] == "table":
            parts.append(block["html"])
        else:
            parts.append(f'<div class="hwpx-line">{block["html"]}</div>')
    return "".join(parts)


def split_hwpx_runs(runs, offset):
    before, after = [], []
    consumed = 0
    for raw_text, props in runs:
        length = len(raw_text)
        if consumed + length <= offset:
            before.append((raw_text, props))
        elif consumed >= offset:
            after.append((raw_text, props))
        else:
            cut = offset - consumed
            before.append((raw_text[:cut], props))
            after.append((raw_text[cut:], props))
        consumed += length
    return before, after


def extract_hwpx_sections(filepath):
    blocks = extract_hwpx_blocks(filepath)

    article_pattern = re.compile(r'^제\s*\d+\s*조(?:의\s*\d+)?(?:\s*\([^)]*\))?')
    simple_pattern = re.compile(r'^\d+\.\s*\S')

    doc_title_blocks = []
    sections = []
    current_heading = None
    current_body = []

    for block in blocks:
        article_match = article_pattern.match(block["plain"]) if block["type"] == "text" else None
        simple_match = None
        if not article_match and block["type"] == "text":
            simple_match = simple_pattern.match(block["plain"])

        if article_match:
            if current_heading is not None:
                sections.append((current_heading, current_body))

            heading_runs, body_runs = split_hwpx_runs(block["runs"], article_match.end())
            heading_plain, heading_html = render_hwpx_runs(heading_runs)
            current_heading = {"type": "text", "plain": heading_plain, "html": heading_html}

            current_body = []
            body_plain, body_html = render_hwpx_runs(body_runs)
            if body_plain.strip():
                current_body.append({"type": "text", "plain": body_plain, "html": body_html})

        elif simple_match:
            if current_heading is not None:
                sections.append((current_heading, current_body))
            current_heading = block
            current_body = []

        else:
            if current_heading is not None:
                current_body.append(block)
            else:
                doc_title_blocks.append(block)

    if current_heading is not None:
        sections.append((current_heading, current_body))

    doc_title_plain = doc_title_blocks[0]["plain"] if doc_title_blocks else ""
    doc_title_extra_blocks = doc_title_blocks[1:]

    return doc_title_plain, doc_title_extra_blocks, sections


def build_hwpx_doc(filepath, filename):
    # 검색어가 제목/조항 어디에도 안 걸릴 때, 문단 단위로 한 번 더 찾아보기 위한
    # "전체 문단 목록" (app.py의 search_hwpx 마지막 fallback과 동일한 역할)
    all_blocks = extract_hwpx_blocks(filepath)
    all_blocks_simple = [{"plain": b["plain"], "html": b["html"]} for b in all_blocks]

    doc_title_plain, doc_title_extra_blocks, sections = extract_hwpx_sections(filepath)

    # 분류별 둘러보기에서 문서 전체를 보여줄 때 쓸 HTML
    full_parts = [f'<div class="hwpx-heading" style="font-size:24px">{hwpx_escape(doc_title_plain)}</div>']
    full_parts.append(render_hwpx_blocks_html(doc_title_extra_blocks))
    section_list = []
    for heading_block, body_blocks in sections:
        full_parts.append(f'<div class="hwpx-heading">{heading_block["html"]}</div>')
        full_parts.append(render_hwpx_blocks_html(body_blocks))

        body_plain_parts = [b["plain"] for b in body_blocks if b["type"] != "table"]
        body_plain_parts += [b["plain"] for b in body_blocks if b["type"] == "table"]
        section_list.append({
            "headingPlain": heading_block["plain"],
            "headingHtml": heading_block["html"],
            "bodyHtml": render_hwpx_blocks_html(body_blocks),
            "bodyPlain": " ".join(body_plain_parts)
        })

    doc_title_extra_html = render_hwpx_blocks_html(doc_title_extra_blocks)

    return {
        "type": "hwpx",
        "docTitlePlain": doc_title_plain,
        "docTitleExtraHtml": doc_title_extra_html,
        "sections": section_list,
        "allBlocks": all_blocks_simple,
        "fullHtml": "".join(full_parts)
    }


# -----------------------------------------------------
# 전체 변환 실행
# -----------------------------------------------------
def main():
    docs = []

    if not os.path.isdir(UPLOAD_FOLDER):
        print(f"'{UPLOAD_FOLDER}' 폴더를 찾을 수 없습니다. uploads 폴더에 문서를 넣어주세요.")
        return

    filenames = sorted(os.listdir(UPLOAD_FOLDER))

    for filename in filenames:
        lower_name = filename.lower()
        ext = lower_name.rsplit(".", 1)[-1] if "." in lower_name else ""
        if ext not in ALLOWED_EXTENSIONS:
            continue

        filepath = os.path.join(UPLOAD_FOLDER, filename)
        print(f"처리 중: {filename}")

        try:
            if ext == "txt":
                body = build_txt_doc(filepath, filename)
            elif ext == "pdf":
                body = build_pdf_doc(filepath, filename)
            elif ext == "xlsx":
                body = build_xlsx_doc(filepath, filename)
            elif ext == "hwpx":
                body = build_hwpx_doc(filepath, filename)
            else:
                continue
        except Exception as e:
            print(f"  -> 처리 실패: {e}")
            continue

        doc = {
            "filename": filename,
            "title": no_ext(filename),
            "category": get_category_name(filename),
            "categoryPriority": get_category_priority(filename),
            "docPriority": get_doc_priority(filename),
        }
        doc.update(body)
        docs.append(doc)

    # 분류 우선순위, 같은 분류 안 순서로 정렬
    docs.sort(key=lambda d: (d["categoryPriority"], d["docPriority"]))

    with open(OUTPUT_DATA_FILE, "w", encoding="utf-8") as f:
        f.write("// 이 파일은 build_data.py가 자동으로 만든 파일입니다. 직접 수정하지 마세요.\n")
        f.write("const DOCS = ")
        f.write(json.dumps(docs, ensure_ascii=False))
        f.write(";\n")

    print(f"\n완료! {len(docs)}개 문서를 {OUTPUT_DATA_FILE} 에 저장했습니다.")


if __name__ == "__main__":
    main()
