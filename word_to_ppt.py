#!/usr/bin/env python3
"""Word OMML 公式 → PowerPoint 可编辑批量转写。

最简用法（只需 Word + 输出路径 + 每页题数）：
    python3 word_to_ppt.py --docx input.docx --out output.pptx --per-page 2

可选自定义样式模板：
    python3 word_to_ppt.py --docx input.docx --out output.pptx --pptx my-template.pptx
"""

from __future__ import annotations

import argparse
import copy
import math
import re
import shutil
import tempfile
import zipfile
from pathlib import Path
from xml.etree import ElementTree as ET

W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
M_NS = "http://schemas.openxmlformats.org/officeDocument/2006/math"
P_NS = "http://schemas.openxmlformats.org/presentationml/2006/main"
A_NS = "http://schemas.openxmlformats.org/drawingml/2006/main"
A14_NS = "http://schemas.microsoft.com/office/drawing/2010/main"
R_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
REL_NS = "http://schemas.openxmlformats.org/package/2006/relationships"
CT_NS = "http://schemas.openxmlformats.org/package/2006/content-types"
EP_NS = "http://schemas.openxmlformats.org/officeDocument/2006/extended-properties"
SLIDE_CT = "application/vnd.openxmlformats-officedocument.presentationml.slide+xml"

ET.register_namespace("p", P_NS)
ET.register_namespace("a", A_NS)
ET.register_namespace("r", R_NS)
ET.register_namespace("m", M_NS)
ET.register_namespace("a14", A14_NS)
ET.register_namespace("w", W_NS)

DEFAULT_TEMPLATE = Path(__file__).resolve().parent / "assets/default-template.pptx"
LABEL_RE = re.compile(r"^[（(](\d+)[)）]$")
EQ_GAP_X = 80000
DEFAULT_EQ_WIDTH = 3200000
DEFAULT_EQ_HEIGHT = 2100000

# 通用 OMML 格式（Office 可编辑公式标准：Cambria Math + 固定字号 + plain 样式）
MATH_FONT_NAME = "Cambria Math"
MATH_FONT_PT = 14  # 磅，可通过 --math-pt 调整


def _math_font_sz_half_points(pt: int | None = None) -> str:
    # Word w:sz 单位是半磅，14pt → 28
    return str(int((pt or MATH_FONT_PT) * 2))


def _standard_w_rpr(pt: int | None = None) -> ET.Element:
    rpr = ET.Element(f"{{{W_NS}}}rPr")
    fonts = ET.SubElement(rpr, f"{{{W_NS}}}rFonts")
    fonts.set(f"{{{W_NS}}}ascii", MATH_FONT_NAME)
    fonts.set(f"{{{W_NS}}}hAnsi", MATH_FONT_NAME)
    fonts.set(f"{{{W_NS}}}cs", MATH_FONT_NAME)
    fonts.set(f"{{{W_NS}}}eastAsia", MATH_FONT_NAME)
    sz = ET.SubElement(rpr, f"{{{W_NS}}}sz")
    sz.set(f"{{{W_NS}}}val", _math_font_sz_half_points(pt))
    sz_cs = ET.SubElement(rpr, f"{{{W_NS}}}szCs")
    sz_cs.set(f"{{{W_NS}}}val", _math_font_sz_half_points(pt))
    return rpr


def _standard_m_rpr() -> ET.Element:
    rpr = ET.Element(f"{{{M_NS}}}rPr")
    sty = ET.SubElement(rpr, f"{{{M_NS}}}sty")
    sty.set(f"{{{M_NS}}}val", "p")
    return rpr


def _set_w_rpr(parent: ET.Element, pt: int | None = None) -> None:
    for old in parent.findall(f"{{{W_NS}}}rPr"):
        parent.remove(old)
    parent.append(_standard_w_rpr(pt))


def normalize_omml(omath: ET.Element, pt: int | None = None) -> ET.Element:
    """统一 OMML 格式：所有公式走同一套 Cambria Math + 字号 + plain 分式。"""
    root = copy.deepcopy(omath)

    for mr in root.findall(f".//{{{M_NS}}}r"):
        for old in mr.findall(f"{{{M_NS}}}rPr"):
            mr.remove(old)
        mr.insert(0, _standard_m_rpr())
        _set_w_rpr(mr, pt)

    for ctrl in root.findall(f".//{{{M_NS}}}ctrlPr"):
        _set_w_rpr(ctrl, pt)

    for fpr in root.findall(f".//{{{M_NS}}}fPr"):
        typ = fpr.find(f"{{{M_NS}}}type")
        if typ is None:
            typ = ET.SubElement(fpr, f"{{{M_NS}}}type")
        typ.set(f"{{{M_NS}}}val", "bar")
        ctrl = fpr.find(f"{{{M_NS}}}ctrlPr")
        if ctrl is None:
            ctrl = ET.SubElement(fpr, f"{{{M_NS}}}ctrlPr")
        _set_w_rpr(ctrl, pt)

    return root

# A4 竖版（与内置模板一致）
SLIDE_W = 6858000
SLIDE_H = 9903460


def build_computed_slots(per_page: int) -> list[dict]:
    """按每页题数自动计算公式位置（无需模板标签）。"""
    margin_x = 900000
    margin_top = 2200000
    margin_bottom = 600000
    label_w = 550000
    cols = 1 if per_page <= 3 else 2
    rows = math.ceil(per_page / cols)
    content_h = SLIDE_H - margin_top - margin_bottom
    content_w = SLIDE_W - 2 * margin_x
    cell_h = content_h // rows
    cell_w = content_w // cols

    slots: list[dict] = []
    for i in range(per_page):
        r, c = divmod(i, cols)
        label_left = margin_x + c * cell_w
        label_top = margin_top + r * cell_h
        slots.append(
            {
                "index": i + 1,
                "left": label_left + label_w + EQ_GAP_X,
                "top": label_top,
                "width": min(DEFAULT_EQ_WIDTH, cell_w - label_w - EQ_GAP_X - 80000),
                "height": max(DEFAULT_EQ_HEIGHT, cell_h - 100000),
            }
        )
    return slots


def extract_template_slots(pptx_path: Path) -> list[dict]:
    """从模板 slideMaster 读取 （1）（2）… 槽位。"""
    slots: dict[int, dict] = {}
    parts = ["ppt/slideMasters/slideMaster1.xml", "ppt/slideLayouts/slideLayout1.xml"]
    with zipfile.ZipFile(pptx_path) as zf:
        parts.extend(
            sorted(n for n in zf.namelist() if re.match(r"ppt/slides/slide\d+\.xml$", n))
        )
        for part in parts:
            if part not in zf.namelist():
                continue
            root = ET.fromstring(zf.read(part))
            for sp in root.findall(f".//{{{P_NS}}}sp"):
                tx = sp.find(f"{{{P_NS}}}txBody")
                if tx is None:
                    continue
                text = "".join(t.text or "" for t in tx.findall(f".//{{{A_NS}}}t")).strip()
                m = LABEL_RE.match(text)
                if not m:
                    continue
                off = sp.find(f"{{{P_NS}}}spPr/{{{A_NS}}}xfrm/{{{A_NS}}}off")
                ext = sp.find(f"{{{P_NS}}}spPr/{{{A_NS}}}xfrm/{{{A_NS}}}ext")
                if off is None or ext is None:
                    continue
                idx = int(m.group(1))
                slots[idx] = {
                    "index": idx,
                    "label_left": int(off.get("x", 0)),
                    "label_top": int(off.get("y", 0)),
                    "label_width": int(ext.get("cx", 0)),
                    "label_height": int(ext.get("cy", 0)),
                }

    if not slots:
        raise ValueError("模板未找到 （1）（2）… 题号标签")

    ordered = [slots[i] for i in sorted(slots)]
    for slot in ordered:
        slot["left"] = slot["label_left"] + slot["label_width"] + EQ_GAP_X
        slot["top"] = slot["label_top"]
        slot["width"] = DEFAULT_EQ_WIDTH
        slot["height"] = max(slot["label_height"], DEFAULT_EQ_HEIGHT)
    return ordered


def resolve_slots(pptx_path: Path, per_page: int | None) -> list[dict]:
    if per_page is None:
        return extract_template_slots(pptx_path)
    if per_page < 1:
        raise ValueError("--per-page 至少为 1")
    try:
        template_slots = extract_template_slots(pptx_path)
        if len(template_slots) == per_page:
            return template_slots
    except ValueError:
        pass
    return build_computed_slots(per_page)


def _sample_math_run_props(paragraph: ET.Element) -> tuple[ET.Element | None, ET.Element | None]:
    """从段落首个 oMath 取字体属性，供合并时的 plain text 复用。"""
    for om in paragraph.findall(f".//{{{M_NS}}}oMath"):
        for mr in om.findall(f".//{{{M_NS}}}r"):
            return mr.find(f"{{{M_NS}}}rPr"), mr.find(f".//{{{W_NS}}}rPr")
    return None, None


def _append_math_text_run(merged: ET.Element, text: str, m_r_pr: ET.Element | None, w_r_pr: ET.Element | None) -> None:
    mr = ET.SubElement(merged, f"{{{M_NS}}}r")
    if m_r_pr is not None:
        mr.append(copy.deepcopy(m_r_pr))
    if w_r_pr is not None:
        wr = ET.SubElement(mr, f"{{{W_NS}}}rPr")
        for child in w_r_pr:
            wr.append(copy.deepcopy(child))
    mt = ET.SubElement(mr, f"{{{M_NS}}}t")
    mt.text = text


def _paragraph_plain_math_text(run: ET.Element) -> str:
    text = "".join(t.text or "" for t in run.findall(f".//{{{W_NS}}}t"))
    text = re.sub(r"^\(\d+\)", "", text)
    text = text.strip()
    if text in {"", "；", "．", ".", ";", "，", ","}:
        return ""
    return text


SUB_QUESTION_RE = re.compile(r"^\(\d+\)")
MAIN_QUESTION_RE = re.compile(r"^\d+[．.]")


def merge_paragraph_omml(paragraph: ET.Element, text: str) -> ET.Element | None:
    """合并段落内 oMath；小题 (1)(2) 保留 = 等文字，大题题干行只取 oMath。"""
    has_omml = any(
        child.tag in (f"{{{M_NS}}}oMath", f"{{{M_NS}}}oMathPara")
        for child in paragraph
    )
    if not has_omml:
        return None

    include_text_runs = bool(SUB_QUESTION_RE.match(text))
    merged = ET.Element(f"{{{M_NS}}}oMath")
    has_content = False
    sample_m_r_pr, sample_w_r_pr = _sample_math_run_props(paragraph)

    for child in paragraph:
        if child.tag == f"{{{W_NS}}}r":
            if not include_text_runs:
                continue
            plain = _paragraph_plain_math_text(child)
            if not plain:
                continue
            _append_math_text_run(merged, plain, sample_m_r_pr, sample_w_r_pr)
            has_content = True
        elif child.tag == f"{{{M_NS}}}oMath":
            for sub in child:
                merged.append(copy.deepcopy(sub))
            has_content = True
        elif child.tag == f"{{{M_NS}}}oMathPara":
            for om in child.findall(f"{{{M_NS}}}oMath"):
                for sub in om:
                    merged.append(copy.deepcopy(sub))
                has_content = True

    return merged if has_content else None


def _should_extract_paragraph(text: str, paragraph: ET.Element) -> bool:
    if not any(
        child.tag in (f"{{{M_NS}}}oMath", f"{{{M_NS}}}oMathPara")
        for child in paragraph
    ):
        return False
    if SUB_QUESTION_RE.match(text):
        return True
    if MAIN_QUESTION_RE.match(text):
        return True
    return False


def extract_omml_equations(docx_path: Path) -> list[dict]:
    with zipfile.ZipFile(docx_path) as zf:
        root = ET.fromstring(zf.read("word/document.xml"))
    body = root.find(f"{{{W_NS}}}body")
    if body is None:
        raise ValueError("docx 缺少 w:body")

    equations: list[dict] = []
    current_q: int | None = None
    eq_idx = 0

    for p in body.findall(f"{{{W_NS}}}p"):
        text = "".join(t.text or "" for t in p.findall(f".//{{{W_NS}}}t")).strip()
        m_q = re.match(r"^(\d+)[．.]", text)
        if m_q:
            current_q = int(m_q.group(1))

        merged = merge_paragraph_omml(p, text)
        if merged is None or not _should_extract_paragraph(text, p):
            continue

        eq_idx += 1
        sub_m = re.match(r"^\((\d+)\)", text)
        equations.append(
            {
                "id": eq_idx,
                "question": current_q,
                "sub": int(sub_m.group(1)) if sub_m else None,
                "type": "omml",
                "text": text[:40],
                "omml": ET.tostring(merged, encoding="unicode"),
            }
        )

    return equations


def build_layout(num_equations: int, slots: list[dict]) -> dict[int, dict]:
    per_slide = len(slots)
    layout: dict[int, dict] = {}
    for i in range(1, num_equations + 1):
        slide = (i - 1) // per_slide
        slot = slots[(i - 1) % per_slide]
        layout[i] = {
            "slide": slide,
            "left": slot["left"],
            "top": slot["top"],
            "width": slot["width"],
            "height": slot["height"],
            "shape_name": f"Eq{i:03d}",
        }
    return layout


def _next_shape_id(slide_root: ET.Element) -> int:
    ids = []
    for sp in slide_root.findall(f".//{{{P_NS}}}sp"):
        cnv = sp.find(f"{{{P_NS}}}nvSpPr/{{{P_NS}}}cNvPr")
        if cnv is not None and str(cnv.get("id", "")).isdigit():
            ids.append(int(cnv.get("id")))
    return max(ids, default=2000) + 1


def _omml_inner(omml_xml: str) -> ET.Element:
    root = ET.fromstring(omml_xml)
    if root.tag == f"{{{M_NS}}}oMath":
        return root
    if root.tag == f"{{{M_NS}}}oMathPara":
        for child in root:
            if child.tag == f"{{{M_NS}}}oMath":
                return child
    return root


def _format_for_ppt(omml_xml: str, pt: int | None = None, target: str = "office") -> ET.Element:
    """office: a14:m 内 oMathPara+jc=left；wps: a14:m 内 inline oMath（WPS 兼容更好）。"""
    omath = normalize_omml(_omml_inner(omml_xml), pt)
    if target == "wps":
        return omath
    para = ET.Element(f"{{{M_NS}}}oMathPara")
    para_pr = ET.SubElement(para, f"{{{M_NS}}}oMathParaPr")
    jc = ET.SubElement(para_pr, f"{{{M_NS}}}jc")
    jc.set(f"{{{M_NS}}}val", "left")
    para.append(omath)
    return para


def _make_omml_shape(
    shape_id: int,
    name: str,
    box: dict,
    omml_xml: str,
    math_pt: int | None = None,
    target: str = "office",
) -> ET.Element:
    sp = ET.Element(f"{{{P_NS}}}sp")
    nv = ET.SubElement(sp, f"{{{P_NS}}}nvSpPr")
    cnv = ET.SubElement(nv, f"{{{P_NS}}}cNvPr")
    cnv.set("id", str(shape_id))
    cnv.set("name", name)
    ET.SubElement(nv, f"{{{P_NS}}}cNvSpPr")
    ET.SubElement(nv, f"{{{P_NS}}}nvPr")

    spPr = ET.SubElement(sp, f"{{{P_NS}}}spPr")
    xfrm = ET.SubElement(spPr, f"{{{A_NS}}}xfrm")
    off = ET.SubElement(xfrm, f"{{{A_NS}}}off")
    off.set("x", str(box["left"]))
    off.set("y", str(box["top"]))
    ext = ET.SubElement(xfrm, f"{{{A_NS}}}ext")
    ext.set("cx", str(box["width"]))
    ext.set("cy", str(box["height"]))
    prst = ET.SubElement(spPr, f"{{{A_NS}}}prstGeom")
    prst.set("prst", "rect")
    ET.SubElement(prst, f"{{{A_NS}}}avLst")
    ET.SubElement(spPr, f"{{{A_NS}}}noFill")

    tx = ET.SubElement(sp, f"{{{P_NS}}}txBody")
    bodyPr = ET.SubElement(tx, f"{{{A_NS}}}bodyPr")
    bodyPr.set("wrap", "none")
    bodyPr.set("anchor", "t")
    bodyPr.set("lIns", "0")
    bodyPr.set("tIns", "0")
    bodyPr.set("rIns", "0")
    bodyPr.set("bIns", "0")
    bodyPr.set("rtlCol", "0")
    ET.SubElement(bodyPr, f"{{{A_NS}}}noAutofit")
    ET.SubElement(tx, f"{{{A_NS}}}lstStyle")
    ap = ET.SubElement(tx, f"{{{A_NS}}}p")
    pPr = ET.SubElement(ap, f"{{{A_NS}}}pPr")
    pPr.set("algn", "l")
    pPr.set("marL", "0")
    pPr.set("indent", "0")
    wrapper = ET.SubElement(ap, f"{{{A14_NS}}}m")
    wrapper.append(_format_for_ppt(omml_xml, math_pt, target))
    return sp


def _ensure_slides(tmp: Path, slide_count: int) -> None:
    slides_dir = tmp / "ppt/slides"
    rels_dir = slides_dir / "_rels"
    slide1 = slides_dir / "slide1.xml"
    rel1 = rels_dir / "slide1.xml.rels"
    pres_rels_path = tmp / "ppt/_rels/presentation.xml.rels"
    pres_path = tmp / "ppt/presentation.xml"

    pres_rels = ET.parse(pres_rels_path).getroot()
    pres = ET.parse(pres_path).getroot()

    existing_rids = [rel.get("Id") for rel in pres_rels.findall(f"{{{REL_NS}}}Relationship")]
    max_rid_num = max(int(r[3:]) for r in existing_rids if r and r.startswith("rId"))

    sld_id_lst = pres.find(f"{{{P_NS}}}sldIdLst")
    existing_sld_ids = [int(sld.get("id")) for sld in sld_id_lst.findall(f"{{{P_NS}}}sldId")]
    next_sld_id = max(existing_sld_ids, default=255) + 1

    for n in range(2, slide_count + 1):
        shutil.copy2(slide1, slides_dir / f"slide{n}.xml")
        shutil.copy2(rel1, rels_dir / f"slide{n}.xml.rels")
        max_rid_num += 1
        rid = f"rId{max_rid_num}"
        rel = ET.SubElement(pres_rels, f"{{{REL_NS}}}Relationship")
        rel.set("Id", rid)
        rel.set("Type", "http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide")
        rel.set("Target", f"slides/slide{n}.xml")
        sld = ET.SubElement(sld_id_lst, f"{{{P_NS}}}sldId")
        sld.set("id", str(next_sld_id))
        sld.set(f"{{{R_NS}}}id", rid)
        next_sld_id += 1

    pres_rels_path.write_bytes(ET.tostring(pres_rels, encoding="utf-8", xml_declaration=True))
    pres_path.write_bytes(ET.tostring(pres, encoding="utf-8", xml_declaration=True))
    _update_content_types(tmp, slide_count)
    _update_app_slide_count(tmp, slide_count)


def _update_content_types(tmp: Path, slide_count: int) -> None:
    ct_path = tmp / "[Content_Types].xml"
    root = ET.parse(ct_path).getroot()
    existing = {el.get("PartName") for el in root.findall(f"{{{CT_NS}}}Override")}
    for n in range(1, slide_count + 1):
        part = f"/ppt/slides/slide{n}.xml"
        if part in existing:
            continue
        ov = ET.SubElement(root, f"{{{CT_NS}}}Override")
        ov.set("PartName", part)
        ov.set("ContentType", SLIDE_CT)
    ct_path.write_bytes(ET.tostring(root, encoding="utf-8", xml_declaration=True))


def _update_app_slide_count(tmp: Path, slide_count: int) -> None:
    app_path = tmp / "docProps/app.xml"
    if not app_path.exists():
        return
    root = ET.parse(app_path).getroot()
    slides_el = root.find(f"{{{EP_NS}}}Slides")
    if slides_el is not None:
        slides_el.text = str(slide_count)
    app_path.write_bytes(ET.tostring(root, encoding="utf-8", xml_declaration=True))


def inject_pptx(
    pptx_in: Path,
    pptx_out: Path,
    equations: list[dict],
    layout: dict[int, dict],
    math_pt: int | None = None,
    target: str = "office",
) -> None:
    slide_count = max(layout[i]["slide"] for i in layout) + 1
    tmp = Path(tempfile.mkdtemp())
    try:
        with zipfile.ZipFile(pptx_in, "r") as zf:
            zf.extractall(tmp)

        _ensure_slides(tmp, slide_count)

        eq_map = {eq["id"]: eq for eq in equations}
        for slide_idx in range(slide_count):
            sf = tmp / f"ppt/slides/slide{slide_idx + 1}.xml"
            root = ET.fromstring(sf.read_bytes())
            cSld = root.find(f"{{{P_NS}}}cSld")
            sp_tree = cSld.find(f"{{{P_NS}}}spTree") if cSld is not None else None
            if sp_tree is None:
                continue
            for eq_id, box in layout.items():
                if box["slide"] != slide_idx:
                    continue
                eq = eq_map.get(eq_id)
                if not eq or eq["type"] != "omml":
                    continue
                sid = _next_shape_id(root)
                sp = _make_omml_shape(sid, box["shape_name"], box, eq["omml"], math_pt, target)
                sp_tree.append(sp)
            sf.write_bytes(ET.tostring(root, encoding="utf-8", xml_declaration=True))

        out_zip = pptx_out.with_suffix(".zip")
        if out_zip.exists():
            out_zip.unlink()
        shutil.make_archive(str(pptx_out.with_suffix("")), "zip", tmp)
        shutil.move(str(out_zip), pptx_out)
    finally:
        shutil.rmtree(tmp, ignore_errors=True)


def main() -> None:
    ap = argparse.ArgumentParser(description="Word OMML 公式批量转 PowerPoint（可编辑）")
    ap.add_argument("--docx", required=True, type=Path, help="源 Word 文件")
    ap.add_argument("--out", required=True, type=Path, help="输出 pptx 路径")
    ap.add_argument(
        "--per-page",
        type=int,
        default=None,
        help="每页题数（如 2、6）。指定后无需关心模板槽位",
    )
    ap.add_argument(
        "--pptx",
        type=Path,
        default=None,
        help="可选自定义样式模板；默认用内置 assets/default-template.pptx",
    )
    ap.add_argument(
        "--math-pt",
        type=int,
        default=MATH_FONT_PT,
        help=f"统一公式字号（磅），默认 {MATH_FONT_PT}",
    )
    ap.add_argument(
        "--target",
        choices=("office", "wps", "both"),
        default="both",
        help="office=Microsoft PowerPoint；wps=WPS 演示；both=同时输出两份（默认）",
    )
    args = ap.parse_args()

    pptx = args.pptx or DEFAULT_TEMPLATE
    if not pptx.exists():
        raise SystemExit(f"模板不存在: {pptx}")

    equations = extract_omml_equations(args.docx)
    if not equations:
        raise SystemExit("未提取到 OMML 公式，请检查 Word 是否为公式编辑器对象")

    slots = resolve_slots(pptx, args.per_page)
    layout = build_layout(len(equations), slots)
    slides_needed = max(box["slide"] for box in layout.values()) + 1
    per_slide = len(slots)

    outputs: list[tuple[str, Path]] = []
    if args.target in ("office", "both"):
        outputs.append(("office", args.out))
    if args.target == "wps":
        outputs.append(("wps", args.out))
    elif args.target == "both":
        outputs.append(("wps", args.out.with_name(f"{args.out.stem}-wps{args.out.suffix}")))

    for target, out_path in outputs:
        inject_pptx(pptx, out_path, equations, layout, args.math_pt, target)

    print(f"✓ 提取 OMML 公式: {len(equations)} 个")
    print(f"✓ 统一格式: Cambria Math {args.math_pt}pt")
    print(f"✓ 每页题数: {per_slide}")
    print(f"✓ 生成幻灯片: {slides_needed} 页")
    for target, out_path in outputs:
        label = "Microsoft PowerPoint" if target == "office" else "WPS 演示"
        print(f"✓ [{label}] {out_path}")
    if args.target == "both":
        print("  WPS 用户请打开 *-wps.pptx；Office 用户请打开主文件")


if __name__ == "__main__":
    main()
