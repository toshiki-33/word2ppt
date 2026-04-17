"""
Word原稿→Claude.ai構成Markdown→PPTX 変換スクリプト

使い方:
    python convert.py input.md output.pptx
    python convert.py input.md output.pptx --template template/base_template.pptx

入力MDの想定フォーマット（Claude.ai Project が v3 インストラクションで出力するもの）:

    ---
    ## slide1: タイトル
    ### 本文（黒・太字）
    [本文]

    ### 強調（赤 #FF0000・太字）
    - 更新世
    - 完新世

    ### 黄色字＝穴埋めの答え（#FFFF00・太字・上から順アニメーション）
    1. 氷期
    2. 間氷期
    ---
"""
import argparse
import re
from pathlib import Path

from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.enum.text import MSO_ANCHOR, MSO_AUTO_SIZE, PP_ALIGN
from pptx.util import Cm, Pt


FONT_NAME = "游ゴシック"
TITLE_FONT_PT = 60
BODY_FONT_PT_DEFAULT = 38
BODY_FONT_PT_NOTE = 32      # ※で始まる行
BODY_FONT_PT_LONG = 36      # 長文行（45字超）
RED_FONT_PT = 36
YELLOW_FONT_PT_DEFAULT = 44
YELLOW_FONT_PT_LARGE = 48   # 特に強調したい用語
YELLOW_FONT_PT_SMALL = 40   # 短いもの

SLIDE_W_CM = 33.87
SLIDE_H_CM = 19.05

TITLE_BODY_X = 2.3
TITLE_BODY_Y = 3.6
TITLE_BODY_W = 29.2
TITLE_BODY_H = 14.3

BODY_X = 0.3
BODY_Y = 0.2
BODY_W = 33.2
BODY_H = 18.6

RED_X = 0.3
RED_Y = 17.5
RED_W = 33.2
RED_H = 1.3

YELLOW_X = 24.0
YELLOW_Y_START = 0.5
YELLOW_W = 9.5
YELLOW_H = 2.0
YELLOW_V_SPACING = 2.3
YELLOW_COL_W = 4.5


def parse_md(md_text):
    """構成MDをパースしてスライドのリストを返す"""
    slides = []
    blocks = re.split(r"^---\s*$", md_text, flags=re.MULTILINE)
    for block in blocks:
        block = block.strip()
        if not block:
            continue
        m = re.search(r"##\s*slide\s*(\d+)\s*[:：]?\s*(.*)", block, re.IGNORECASE)
        if not m:
            continue
        slide_no = int(m.group(1))
        slide_title = m.group(2).strip()

        body_lines = []
        red_items = []
        yellow_items = []
        current = None

        for raw in block.split("\n"):
            line = raw.rstrip()
            if re.match(r"^\s*###\s*本文", line):
                current = "body"
                continue
            if re.match(r"^\s*###\s*(強調|赤字)", line):
                current = "red"
                continue
            if re.match(r"^\s*###\s*黄色字", line):
                current = "yellow"
                continue
            if re.match(r"^\s*##[^#]", line) or re.match(r"^\s*#[^#]", line):
                continue

            if current == "body":
                stripped = line.strip()
                if stripped.startswith("[") and stripped.endswith("]"):
                    continue
                body_lines.append(line)
            elif current == "red":
                m2 = re.match(r"^\s*[-・*]\s*(.+)$", line)
                if m2:
                    red_items.append(m2.group(1).strip())
            elif current == "yellow":
                m2 = re.match(r"^\s*\d+\.\s*(.+)$", line)
                if m2:
                    txt = m2.group(1).strip()
                    txt = re.sub(r"^【要確認】\s*", "【要確認】", txt)
                    yellow_items.append(txt)

        while body_lines and not body_lines[0].strip():
            body_lines.pop(0)
        while body_lines and not body_lines[-1].strip():
            body_lines.pop()

        slides.append(
            {
                "no": slide_no,
                "title": slide_title,
                "body": "\n".join(body_lines),
                "red": red_items,
                "yellow": yellow_items,
            }
        )
    slides.sort(key=lambda s: s["no"])
    return slides


def remove_all_slides(prs):
    """テンプレの既存スライドを package ごと削除する"""
    from pptx.oxml.ns import qn

    sldIdLst = prs.slides._sldIdLst
    slide_id_elements = list(sldIdLst)
    for sldId in slide_id_elements:
        rId = sldId.get(qn("r:id"))
        prs.part.drop_rel(rId)
        sldIdLst.remove(sldId)

    pkg = prs.part.package
    for part in list(pkg.iter_parts()):
        if part.partname.startswith("/ppt/slides/slide"):
            try:
                pkg._parts.pop(part.partname, None)
            except Exception:
                pass


def set_black_background(slide):
    """スライド背景を黒にする（XML直接操作）"""
    from pptx.oxml.ns import qn
    from lxml import etree

    spTree = slide.shapes._spTree
    cSld = spTree.getparent()
    for existing_bg in cSld.findall(qn("p:bg")):
        cSld.remove(existing_bg)

    bg_xml = (
        '<p:bg xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" '
        'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">'
        "<p:bgPr>"
        '<a:solidFill><a:srgbClr val="000000"/></a:solidFill>'
        '<a:effectLst/>'
        "</p:bgPr>"
        "</p:bg>"
    )
    bg_elem = etree.fromstring(bg_xml)
    cSld.insert(0, bg_elem)


def set_run_common(run, text, pt, color_rgb, bold=True):
    run.text = text
    run.font.size = Pt(pt)
    run.font.bold = bold
    run.font.name = FONT_NAME
    run.font.color.rgb = color_rgb


CHARS_PER_LINE_38PT = 18
MAX_BODY_LINES_PER_SLIDE = 12
MAX_YELLOW_PER_SLIDE = 7


def count_wrapped_lines(text, chars_per_line=CHARS_PER_LINE_38PT):
    """折り返しと空行を考慮した実行行数"""
    total = 0
    for line in text.split("\n"):
        if line.strip():
            total += max(1, (len(line) + chars_per_line - 1) // chars_per_line)
        else:
            total += 1
    return total


def split_body_smart(body, max_lines=MAX_BODY_LINES_PER_SLIDE):
    """本文を意味の区切り優先で分割する"""
    lines = body.split("\n")
    parts = []
    current = []

    def wrap():
        return count_wrapped_lines("\n".join(current))

    for line in lines:
        current.append(line)
        if wrap() > max_lines and len(current) > 1:
            cut = None
            for j in range(len(current) - 2, 0, -1):
                if not current[j].strip():
                    cut = j
                    break
                stripped = current[j].strip()
                if stripped.startswith(("◎", "※", "・")) or (
                    stripped and stripped[0] in "（("
                ):
                    cut = j
                    break
            if cut is None:
                cut = len(current) - 1

            head = current[:cut]
            while head and not head[-1].strip():
                head.pop()
            if head:
                parts.append("\n".join(head).strip())

            tail = current[cut:]
            while tail and not tail[0].strip():
                tail.pop(0)
            current = tail

    if current:
        leftover = "\n".join(current).strip()
        if leftover:
            parts.append(leftover)

    return parts or [body]


def assign_yellow_to_parts(body_parts, yellow_items):
    """黄色字を本文パートの穴埋め番号出現順に割り振る"""
    result = [[] for _ in body_parts]
    y_idx = 0
    for p_idx, part in enumerate(body_parts):
        hole_count = len(re.findall(r"[（(]\s*\d+", part))
        for _ in range(hole_count):
            if y_idx < len(yellow_items):
                result[p_idx].append(yellow_items[y_idx])
                y_idx += 1
    while y_idx < len(yellow_items):
        result[-1].append(yellow_items[y_idx])
        y_idx += 1
    return result


def split_yellow_if_overflow(slides_expanded):
    """黄色字が上限超のスライドを更に半分に分割（本文は同一で繰り返し）"""
    result = []
    for sd in slides_expanded:
        if len(sd["yellow"]) <= MAX_YELLOW_PER_SLIDE:
            result.append(sd)
            continue
        y = sd["yellow"]
        half = (len(y) + 1) // 2
        a = dict(sd, yellow=y[:half])
        b = dict(sd, yellow=y[half:], red=[], title=sd["title"] + "（続き）")
        result.append(a)
        result.append(b)
    return result


def split_slide_if_needed(slide_data):
    """本文の行数または黄色字個数が上限を超えたらスライド分割"""
    body = slide_data["body"]
    yellow = slide_data["yellow"]

    wrapped = count_wrapped_lines(body)
    if wrapped <= MAX_BODY_LINES_PER_SLIDE and len(yellow) <= MAX_YELLOW_PER_SLIDE:
        return [slide_data]

    if wrapped > MAX_BODY_LINES_PER_SLIDE:
        body_parts = split_body_smart(body)
    else:
        body_parts = [body]

    yellow_per_part = assign_yellow_to_parts(body_parts, yellow)

    expanded = []
    for idx, (bp, yp) in enumerate(zip(body_parts, yellow_per_part)):
        suffix = "" if idx == 0 else "（続き）"
        expanded.append(
            {
                "no": slide_data["no"],
                "title": slide_data["title"] + suffix,
                "body": bp,
                "red": slide_data["red"] if idx == 0 else [],
                "yellow": yp,
            }
        )

    return split_yellow_if_overflow(expanded)


def yellow_line_size(text):
    return YELLOW_FONT_PT_DEFAULT


def is_title_slide(slide):
    if slide["no"] == 1:
        return True
    body = slide["body"].strip()
    if not slide["yellow"] and not slide["red"] and len(body) < 80 and "\n" not in body:
        return True
    return False


def add_title_slide(prs, slide_data):
    blank = prs.slide_layouts[6]
    slide = prs.slides.add_slide(blank)
    set_black_background(slide)
    tb = slide.shapes.add_textbox(
        Cm(TITLE_BODY_X), Cm(TITLE_BODY_Y), Cm(TITLE_BODY_W), Cm(TITLE_BODY_H)
    )
    tf = tb.text_frame
    tf.word_wrap = True
    tf.vertical_anchor = MSO_ANCHOR.MIDDLE
    p = tf.paragraphs[0]
    p.alignment = PP_ALIGN.CENTER
    text = slide_data["body"].strip() or slide_data["title"]
    run = p.add_run()
    set_run_common(run, text, TITLE_FONT_PT, RGBColor(0xFF, 0xFF, 0xFF))
    return slide


def add_body_block(slide, lines):
    tb = slide.shapes.add_textbox(Cm(BODY_X), Cm(BODY_Y), Cm(BODY_W), Cm(BODY_H))
    tf = tb.text_frame
    tf.word_wrap = True
    tf.auto_size = MSO_AUTO_SIZE.NONE
    tf.margin_left = Cm(0.1)
    tf.margin_right = Cm(0.1)
    tf.margin_top = Cm(0.05)
    tf.margin_bottom = Cm(0.05)
    first = True
    for raw in lines.split("\n"):
        line = raw.rstrip()
        if first:
            p = tf.paragraphs[0]
            first = False
        else:
            p = tf.add_paragraph()
        p.line_spacing = 1.0
        p.space_before = Pt(0)
        p.space_after = Pt(0)
        if not line.strip():
            continue
        run = p.add_run()
        set_run_common(run, line, BODY_FONT_PT_DEFAULT, RGBColor(0xFF, 0xFF, 0xFF))
    return tb


def add_red_block(slide, red_items):
    if not red_items:
        return None
    tb = slide.shapes.add_textbox(Cm(RED_X), Cm(RED_Y), Cm(RED_W), Cm(RED_H))
    tf = tb.text_frame
    tf.word_wrap = True
    text = "  ".join(red_items)
    p = tf.paragraphs[0]
    run = p.add_run()
    set_run_common(run, text, RED_FONT_PT, RGBColor(0xFF, 0x00, 0x00))
    return tb


def add_yellow_blocks(slide, yellow_items):
    placed = []
    for idx, text in enumerate(yellow_items):
        col = 0
        row = idx
        y = YELLOW_Y_START + row * YELLOW_V_SPACING
        while y + YELLOW_H > SLIDE_H_CM - 0.3:
            col += 1
            row = idx - col * int((SLIDE_H_CM - YELLOW_Y_START) // YELLOW_V_SPACING)
            y = YELLOW_Y_START + row * YELLOW_V_SPACING
            if col > 1:
                break
        x = YELLOW_X + col * (YELLOW_COL_W + 0.2)
        tb = slide.shapes.add_textbox(Cm(x), Cm(y), Cm(YELLOW_W if col == 0 else YELLOW_COL_W), Cm(YELLOW_H))
        tf = tb.text_frame
        tf.word_wrap = True
        tf.vertical_anchor = MSO_ANCHOR.MIDDLE
        p = tf.paragraphs[0]
        p.alignment = PP_ALIGN.LEFT
        pt = yellow_line_size(text)
        numbered_text = f"{idx + 1}. {text}"
        run = p.add_run()
        set_run_common(run, numbered_text, pt, RGBColor(0xFF, 0xFF, 0x00))
        placed.append(tb)
    return placed


def add_content_slide(prs, slide_data):
    blank = prs.slide_layouts[6]
    slide = prs.slides.add_slide(blank)
    set_black_background(slide)
    add_body_block(slide, slide_data["body"])
    add_red_block(slide, slide_data["red"])
    add_yellow_blocks(slide, slide_data["yellow"])
    return slide


def convert_md_to_pptx(md_text, output, template_path=None):
    """構成MD文字列をPPTXに変換する。

    Args:
        md_text (str): 構成Markdown本文
        output (str | BinaryIO): 保存先パス or 書き込み可能なバイナリストリーム
        template_path (str, optional): テンプレpptxパス
    Returns:
        dict: {"original": 元スライド数, "expanded": 分割後スライド数}
    """
    tpl_path = template_path or str(
        Path(__file__).parent / "template" / "base_template.pptx"
    )
    prs = Presentation(tpl_path)
    remove_all_slides(prs)

    slides = parse_md(md_text)
    if not slides:
        raise ValueError("構成MDからスライドを検出できませんでした。フォーマットを確認してください。")

    expanded = []
    for sd in slides:
        if is_title_slide(sd):
            expanded.append(sd)
        else:
            expanded.extend(split_slide_if_needed(sd))

    for sd in expanded:
        if is_title_slide(sd):
            add_title_slide(prs, sd)
        else:
            add_content_slide(prs, sd)

    prs.save(output)
    return {"original": len(slides), "expanded": len(expanded)}


def main():
    ap = argparse.ArgumentParser(description="構成MD→PPTX 変換")
    ap.add_argument("input_md", help="構成Markdownファイル")
    ap.add_argument("output_pptx", help="出力PPTXパス")
    ap.add_argument(
        "--template",
        default=None,
        help="テンプレPPTX（省略時は template/base_template.pptx）",
    )
    args = ap.parse_args()

    md = Path(args.input_md).read_text(encoding="utf-8")
    result = convert_md_to_pptx(md, args.output_pptx, args.template)
    split_added = result["expanded"] - result["original"]
    print(
        f"saved: {args.output_pptx} ({result['expanded']} slides"
        + (f", {split_added}枚は自動分割" if split_added else "")
        + ")"
    )


if __name__ == "__main__":
    main()
