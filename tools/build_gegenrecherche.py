#!/usr/bin/env python3
"""GEGENRECHERCHE 원고(.txt) → 낭독 교재 docx 변환기.

원고 형식:
  @part 1 / @ep 1 / @title ... / @ko ... / @sub ... / @info ...
  @cast NAME | 설명      (여러 줄)
  @note ※ ...           (여러 줄 가능)
  @next ...
  # 장면 제목
  NAME: 독일어 대사
  = 표준어 풀이 (선택, 사투리 대사용)
  > 한국어 대역
사용법: build_gegenrecherche.py <원고.txt> [...]   (턴 수 검사 포함)
"""
import re, sys, pathlib
from docx import Document
from docx.shared import Pt, RGBColor, Twips
from docx.enum.text import WD_ALIGN_PARAGRAPH

ROOT = pathlib.Path(__file__).resolve().parent.parent
OUT = ROOT / "화별_대본" / "GEGENRECHERCHE"
EXPECTED_TURNS = 150

COLORS = {
    "JAN": "0B5394", "HELENE": "1F3864", "MIRA": "A64D79", "TOBI": "AA3939",
    "SWETA": "38761D", "LEON": "B45F06", "ELKE": "7F6000", "KATRIN": "45818E",
    "DANA": "674EA7", "THAL": "434343", "KASPER": "666666", "SPATZ": "990000",
}
FALLBACK = ["2E75B6", "6AA84F", "C27BA0", "E69138", "3D85C6", "8E7CC3", "CC4125", "76A5AF"]

def color_for(name, extra):
    if name in COLORS:
        return COLORS[name]
    if name not in extra:
        extra[name] = FALLBACK[len(extra) % len(FALLBACK)]
    return extra[name]

def para(doc, text="", *, size=None, color=None, bold=False, italic=False,
         before=None, after=None, center=False):
    p = doc.add_paragraph()
    pf = p.paragraph_format
    if before is not None: pf.space_before = Twips(before)
    if after is not None: pf.space_after = Twips(after)
    if center: p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    if text:
        r = p.add_run(text)
        r.bold = bold; r.italic = italic
        if size: r.font.size = Pt(size)
        if color: r.font.color.rgb = RGBColor.from_string(color)
    return p

def parse(path):
    meta = {"cast": [], "note": []}
    body = []  # ("scene", title) | ("turn", name, de, gloss, ko)
    for raw in path.read_text(encoding="utf-8").splitlines():
        line = raw.rstrip()
        if not line.strip():
            continue
        if line.startswith("@"):
            key, _, val = line[1:].partition(" ")
            if key == "cast": meta["cast"].append(val)
            elif key == "note": meta["note"].append(val)
            else: meta[key] = val
        elif line.startswith("# "):
            body.append(["scene", line[2:].strip()])
        elif line.startswith("= "):
            body[-1][3] = line[2:].strip()
        elif line.startswith("> "):
            body[-1][4] = line[2:].strip()
        else:
            m = re.match(r"^([A-ZÄÖÜ][A-ZÄÖÜ0-9 .\-]*):\s+(.*)$", line)
            if not m:
                raise SystemExit(f"{path.name}: 해석할 수 없는 줄: {line!r}")
            body.append(["turn", m.group(1).strip(), m.group(2).strip(), None, None])
    for b in body:
        if b[0] == "turn" and not b[4]:
            raise SystemExit(f"{path.name}: 한국어 대역 없음: {b[1]}: {b[2]}")
    return meta, body

def build(path):
    meta, body = parse(path)
    turns = sum(1 for b in body if b[0] == "turn")
    if turns != EXPECTED_TURNS:
        raise SystemExit(f"{path.name}: 턴 수 {turns} (기대값 {EXPECTED_TURNS})")
    doc = Document()
    st = doc.styles["Normal"]; st.font.size = Pt(11)
    sec = doc.sections[0]
    sec.page_width, sec.page_height = Twips(11906), Twips(16838)
    sec.top_margin = sec.bottom_margin = Twips(1134)
    sec.left_margin = sec.right_margin = Twips(1417)

    para(doc, "GEGENRECHERCHE", size=15, color="808080", bold=True, before=2400, after=200, center=True)
    para(doc, "Berliner Akten · 베를린식 독일어 낭독 교재", size=13, color="808080", before=0, after=100, center=True)
    para(doc, f"{meta['part']}부 · {meta['ep']}화", size=12, color="1F3864", bold=True, before=600, after=100, center=True)
    para(doc, meta["title"], size=20, color="1F3864", bold=True, before=0, after=40, center=True)
    para(doc, meta["sub"], size=12, color="404040", before=600, after=40, center=True)
    para(doc, meta["info"], size=9, color="595959", before=0, after=40, center=True)
    para(doc, "지문 없이 대사만으로 구성 · 독일어 본문 + (사투리 대사는 표준어 풀이) + 한국어 대역",
         size=9, color="595959", before=0, after=100, center=True)
    para(doc, "등장인물", size=10, color="1F3864", bold=True, before=500, after=150, center=True)
    for c in meta["cast"]:
        name, _, desc = c.partition("|")
        p = para(doc, after=40, center=True)
        r = p.add_run(name.strip() + "  "); r.bold = True; r.font.size = Pt(10)
        r = p.add_run(desc.strip()); r.font.size = Pt(9); r.font.color.rgb = RGBColor.from_string("595959")
    for n in meta["note"]:
        para(doc, n, size=8, color="808080", before=120, after=0)
    para(doc)

    extra = {}
    for b in body:
        if b[0] == "scene":
            para(doc, b[1], size=12, color="1F3864", bold=True, before=400, after=240)
            continue
        _, name, de, gloss, ko = b
        p = para(doc, before=160, after=20)
        r = p.add_run(name); r.bold = True; r.font.size = Pt(10)
        r.font.color.rgb = RGBColor.from_string(color_for(name, extra))
        r = p.add_run("   " + de); r.font.size = Pt(11); r.font.color.rgb = RGBColor.from_string("1A1A1A")
        if gloss:
            para(doc, "≈ " + gloss, size=8.5, color="5B7FA6", italic=True, before=0, after=0)
        para(doc, ko, size=9, color="808080", before=0, after=0)

    para(doc, "다음 화 예고", size=12, color="1F3864", bold=True, before=500, after=100)
    para(doc, meta["next"], size=9, color="595959")

    folder = OUT / f"{meta['part']}부"
    folder.mkdir(parents=True, exist_ok=True)
    safe = re.sub(r"[^\w\-]+", "_", meta["title"].replace("ä","ae").replace("ö","oe").replace("ü","ue")
                  .replace("Ä","Ae").replace("Ö","Oe").replace("Ü","Ue").replace("ß","ss")).strip("_")
    out = folder / f"{meta['part']}부_{meta['ep']}화_{safe}.docx"
    doc.save(out)
    print(f"OK {out.relative_to(ROOT)}  ({turns}턴)")

if __name__ == "__main__":
    for a in sys.argv[1:]:
        build(pathlib.Path(a))
