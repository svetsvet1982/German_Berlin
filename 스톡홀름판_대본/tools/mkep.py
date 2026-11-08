# -*- coding: utf-8 -*-
"""원문(txt) → 베를린판과 같은 서식의 docx.  사용: python3 mkep.py 원문.txt 출력폴더
원문 형식:
  PART: 1 / EP: 1 / TITLE: / SUBTITLE: / META: / CAST: NAME | 설명 / NOTE: / NEXT:
  # 장면 제목
  NAME: 스웨덴어 대사
  = 한국어 대역
"""
import sys, os, re
from docx import Document
from docx.shared import Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn

COLORS = {"BIGGAN": "7F6000", "JOHAN": "0B5394", "DANIEL": "AA3939", "MAJA": "38761D",
          "ANTON": "2E75B6", "ELSA": "A64D79", "OXENSTIERNA": "7F7F7F", "WICKSELL": "6D4C41",
          "STRINDBERG": "A0522D", "HAMMARSKJÖLD": "8E24AA", "WÄGNER": "C00000", "NOBEL": "7030A0",
          "LAGERLÖF": "5E35B1", "CHYDENIUS": "455A64", "LINNÉ": "2E7D32", "KEY": "AD1457"}
EXTRA = ["B45F06", "0E7C86", "4A6741", "5B9BD5", "9C6500", "BF8F00", "C55A11", "974706", "767171", "00838F"]

def trans(s):
    for a, b in [("å", "aa"), ("ä", "ae"), ("ö", "oe"), ("Å", "Aa"), ("Ä", "Ae"), ("Ö", "Oe"), ("é", "e"), ("É", "E")]:
        s = s.replace(a, b)
    return re.sub(r"[^A-Za-z0-9]+", "_", s).strip("_")

def parse(path):
    h = {"CAST": []}; body = []; cur = None
    for raw in open(path, encoding="utf-8"):
        line = raw.rstrip("\n").strip()
        if not line: continue
        m = re.match(r"^(PART|EP|TITLE|SUBTITLE|META|NOTE|NEXTHEAD|NEXT|CAST):\s*(.*)$", line)
        if m:
            k, v = m.groups()
            if k == "CAST":
                n, d = [x.strip() for x in v.split("|", 1)]; h["CAST"].append((n, d))
            else: h[k] = v
            continue
        if line.startswith("# "):
            body.append(("H", line[2:].strip())); continue
        if line.startswith("= "):
            assert cur and cur[3] is None, f"대역 위치 오류: {line}"
            cur[3] = line[2:].strip(); continue
        m = re.match(r"^([A-ZÅÄÖÉ\-]+):\s+(.+)$", line)
        assert m, f"해석 불가 줄: {line}"
        cur = ["T", m.group(1), m.group(2), None]; body.append(cur)
    return h, body

def run(p, text, size, color=None, bold=False):
    r = p.add_run(text); r.font.size = Pt(size); r.bold = bold
    if color: r.font.color.rgb = RGBColor.from_string(color)
    return r

def sp(p, before=None, after=None):
    f = p.paragraph_format
    if before is not None: f.space_before = Pt(before)
    if after is not None: f.space_after = Pt(after)

def build(path, outdir):
    h, body = parse(path)
    turns = [b for b in body if b[0] == "T"]
    names = {n for n, _ in h["CAST"]}
    errs = [f"대역 없음: {t[2]}" for t in turns if not t[3]]
    errs += [f"등장인물 목록에 없음: {t[1]}" for t in turns if t[1] not in names]
    if len(turns) != 120: errs.append(f"턴 수 {len(turns)} ≠ 120")
    if errs: sys.exit("\n".join(errs))
    for a, b in zip(turns, turns[1:]):
        if a[1] == b[1]: print("경고: 같은 화자 연속 —", b[1], b[2][:40])
    color = dict(COLORS); i = 0
    for n, _ in h["CAST"]:
        if n not in color: color[n] = EXTRA[i % len(EXTRA)]; i += 1
    d = Document()
    d.styles["Normal"].element.get_or_add_rPr().get_or_add_rFonts().set(qn("w:eastAsia"), "맑은 고딕")
    C = WD_ALIGN_PARAGRAPH.CENTER
    def cp(before, after):
        p = d.add_paragraph(); p.alignment = C; sp(p, before, after); return p
    run(cp(120, 10), "STOCKHOLM", 15, "808080", True)
    run(cp(0, 5), "스웨덴어 낭독 교재", 13, "808080")
    run(cp(30, 5), f"{h['PART']}부 · {h['EP']}화", 12, "1F3864", True)
    run(cp(0, 2), h["TITLE"], 20, "1F3864", True)
    run(cp(30, 2), h["SUBTITLE"], 12, "404040")
    run(cp(0, 2), h["META"], 9, "595959")
    run(cp(0, 5), "지문 없이 대사만으로 구성 · 스웨덴어 본문 + 한국어 대역", 9, "595959")
    run(cp(25, 7.5), "등장인물", 10, "1F3864", True)
    for n, desc in h["CAST"]:
        p = cp(None, 2); run(p, n + "  ", 10, None, True); run(p, desc, 9, "595959")
    p = d.add_paragraph(); sp(p, 0, 0); run(p, h["NOTE"], 8, "808080")
    d.add_paragraph()
    for b in body:
        if b[0] == "H":
            p = d.add_paragraph(); sp(p, 20, 12); run(p, b[1], 12, "1F3864", True)
        else:
            _, n, sv, ko = b
            p = d.add_paragraph(); sp(p, 8, 1); run(p, n, 10, color[n], True); run(p, "   " + sv, 11, "1A1A1A")
            p = d.add_paragraph(); sp(p, None, 0); run(p, ko, 9, "808080")
    p = d.add_paragraph(); sp(p, 25, 5); run(p, h.get("NEXTHEAD", "다음 화 예고"), 12, "1F3864", True)
    run(d.add_paragraph(), h["NEXT"], 9, "595959")
    sub = os.path.join(outdir, f"{h['PART']}부"); os.makedirs(sub, exist_ok=True)
    out = os.path.join(sub, f"{h['PART']}부_{h['EP']}화_{trans(h['TITLE'])}.docx")
    d.save(out); print(out, len(turns), "턴")

if __name__ == "__main__":
    build(sys.argv[1], sys.argv[2])
