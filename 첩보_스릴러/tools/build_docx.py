#!/usr/bin/env python3
"""원고(.txt) → 화별 대본(.docx) 변환기 · 『DER GÄRTNER』

원고 형식
---------
#META
part: 1
episode: 1
title: Der Karton ohne Nummer
subtitle_ko: 번호 없는 상자
place_ko: 장벽 붕괴 기념일, BND 본부 지하 창고        (표지 부제 뒤에 붙는 짧은 장소)
setting: 2026년 11월 9일(월) 오후 다섯 시, BND 본부 지하 기록보관 창고
level: B1
#CAST
JAKOB | BND 기록보관 부서 사무관 · 44세
#INTRO
※ 한국어 도입글 (한 문단)
#SECTION 도입 — DER NEUNTE NOVEMBER
JAKOB: Deutscher Satz.
= 한국어 대역
#NEXT
1부 2화 — «...»(...) — 예고 문장

사용법:  python3 build_docx.py 원고파일.txt [...]   (출력: 첩보_스릴러/화별_대본/n부/)
검사만:  python3 build_docx.py --check 원고파일.txt
"""
import re
import sys
from pathlib import Path

from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn
from docx.shared import Pt, RGBColor, Twips

ROOT = Path(__file__).resolve().parent.parent
OUT_DIR = ROOT / "화별_대본"
TURN_MIN, TURN_MAX = 120, 150

# 화자 색상 (기존 『BERLIN』 시리즈 팔레트 + 확장)
PALETTE = ["0B5394", "AA3939", "38761D", "7F6000", "A64D79", "2E75B6",
           "B45F06", "134F5C", "674EA7", "990000", "45818E", "6AA84F"]
FIXED = {"JAKOB": "0B5394", "KATRIN": "AA3939", "SEIDLITZ": "674EA7",
         "KRANZ": "7F6000", "ROTH": "990000", "MIRA": "38761D",
         "RENATE": "B45F06", "LEA": "A64D79", "AYLIN": "134F5C",
         "BRÜCKNER": "45818E", "FELDMANN": "2E75B6"}
GERMAN_NAME = re.compile(r"^([A-ZÄÖÜ][A-ZÄÖÜ\- ]*[A-ZÄÖÜ]):\s*(.+)$")


class ManuscriptError(Exception):
    pass


def parse(path):
    meta, cast, intro, sections, nxt = {}, [], [], [], []
    block, cur = None, None
    for n, raw in enumerate(Path(path).read_text(encoding="utf-8").splitlines(), 1):
        line = raw.rstrip()
        if not line.strip():
            continue
        if line.startswith("#SECTION"):
            block = "SECTION"
            cur = {"title": line[len("#SECTION"):].strip(), "turns": []}
            sections.append(cur)
            continue
        if line.startswith("#") and line[1:].strip() in ("META", "CAST", "INTRO", "NEXT"):
            block = line[1:].strip()
            continue
        if block == "META":
            k, _, v = line.partition(":")
            meta[k.strip()] = v.strip()
        elif block == "CAST":
            name, _, desc = line.partition("|")
            cast.append((name.strip(), desc.strip()))
        elif block == "INTRO":
            intro.append(line.strip())
        elif block == "NEXT":
            nxt.append(line.strip())
        elif block == "SECTION":
            if line.startswith("="):
                if not cur["turns"] or cur["turns"][-1][2] is not None:
                    raise ManuscriptError(f"{path}:{n}: 대역(=) 앞에 대사가 없음")
                cur["turns"][-1][2] = line[1:].strip()
            else:
                m = GERMAN_NAME.match(line)
                if not m:
                    raise ManuscriptError(f"{path}:{n}: 대사 형식 오류 → {line[:60]}")
                if cur["turns"] and cur["turns"][-1][2] is None:
                    raise ManuscriptError(f"{path}:{n}: 앞 대사에 대역이 없음")
                cur["turns"].append([m.group(1), m.group(2).strip(), None])
        else:
            raise ManuscriptError(f"{path}:{n}: 블록 밖의 텍스트")
    return meta, cast, " ".join(intro), sections, " ".join(nxt)


def validate(path, meta, cast, sections):
    errs = []
    for k in ("part", "episode", "title", "subtitle_ko", "setting", "level"):
        if k not in meta:
            errs.append(f"META에 {k} 없음")
    names = {c[0] for c in cast}
    turns = [t for s in sections for t in s["turns"]]
    for spk, de, ko in turns:
        if spk not in names:
            errs.append(f"등장인물 목록에 없는 화자: {spk}")
        if ko is None:
            errs.append(f"대역 누락: {spk}: {de[:40]}")
    used = {t[0] for t in turns}
    for nm in names - used:
        errs.append(f"등장인물에 있으나 대사 없음: {nm}")
    if not (TURN_MIN <= len(turns) <= TURN_MAX):
        errs.append(f"턴 수 {len(turns)} (허용 {TURN_MIN}~{TURN_MAX})")
    return sorted(set(errs)), len(turns)


def run(p, text, size, bold=False, color=None):
    r = p.add_run(text)
    r.font.size = Pt(size)
    r.bold = bold
    if color:
        r.font.color.rgb = RGBColor.from_string(color)
    return r


def para(doc, before=None, after=None, center=False):
    p = doc.add_paragraph()
    pf = p.paragraph_format
    if before is not None:
        pf.space_before = Twips(before)
    if after is not None:
        pf.space_after = Twips(after)
    if center:
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    return p


def build(path):
    meta, cast, intro, sections, nxt = parse(path)
    errs, nturns = validate(path, meta, cast, sections)
    if errs:
        raise ManuscriptError(f"{path}:\n  - " + "\n  - ".join(errs))

    colors, i = {}, 0
    for name, _ in cast:
        if name in FIXED:
            colors[name] = FIXED[name]
        else:
            while PALETTE[i % len(PALETTE)] in FIXED.values() and i < 50:
                i += 1
            colors[name] = PALETTE[i % len(PALETTE)]
            i += 1

    doc = Document()
    st = doc.styles["Normal"]
    st.font.name = "Calibri"
    st.element.rPr.rFonts.set(qn("w:eastAsia"), "Malgun Gothic")
    sec = doc.sections[0]
    sec.page_width, sec.page_height = Twips(11906), Twips(16838)
    sec.top_margin = sec.bottom_margin = Twips(1134)
    sec.left_margin = sec.right_margin = Twips(1417)

    part, ep = meta["part"], meta["episode"]
    run(para(doc, 2400, 200, True), "DER GÄRTNER", 15, True, "808080")
    run(para(doc, 0, 100, True), "독일어 낭독 교재 · 첩보 스릴러", 13, False, "808080")
    run(para(doc, 600, 100, True), f"{part}부 · {ep}화", 12, True, "1F3864")
    run(para(doc, 0, 40, True), meta["title"], 20, True, "1F3864")
    sub = meta["subtitle_ko"] + (f" · {meta['place_ko']}" if meta.get("place_ko") else "")
    run(para(doc, 600, 40, True), sub, 12, False, "404040")
    run(para(doc, 0, 40, True),
        f"대사 {nturns}턴 · 난이도 {meta['level']} · 배경: {meta['setting']}", 9, False, "595959")
    run(para(doc, 0, 100, True), "지문 없이 대사만으로 구성 · 독일어 본문 + 한국어 대역", 9, False, "595959")
    run(para(doc, 500, 150, True), "등장인물", 10, True, "1F3864")
    for name, desc in cast:
        p = para(doc, after=40, center=True)
        run(p, f"{name}  ", 10, True)
        run(p, desc, 9, False, "595959")
    run(para(doc, 0, 0), intro, 8, False, "808080")
    doc.add_paragraph()

    for s in sections:
        run(para(doc, 400, 240), s["title"], 12, True, "1F3864")
        for spk, de, ko in s["turns"]:
            p = para(doc, 160, 20)
            run(p, spk, 10, True, colors[spk])
            run(p, f"   {de}", 11, False, "1A1A1A")
            run(para(doc, after=0), ko, 9, False, "808080")

    if nxt:
        run(para(doc, 500, 100), "다음 화 예고", 12, True, "1F3864")
        run(doc.add_paragraph(), nxt, 9, False, "595959")

    safe = re.sub(r"[^\w]+", "_", meta["title"].translate(str.maketrans(
        {"ä": "ae", "ö": "oe", "ü": "ue", "Ä": "Ae", "Ö": "Oe", "Ü": "Ue", "ß": "ss"}))).strip("_")
    out = OUT_DIR / f"{part}부" / f"{part}부_{ep}화_{safe}.docx"
    out.parent.mkdir(parents=True, exist_ok=True)
    doc.core_properties.title = f"DER GÄRTNER {part}부 {ep}화 — {meta['title']}"
    doc.save(out)
    return out, nturns


def main(argv):
    check = "--check" in argv
    files = [a for a in argv if a != "--check"]
    bad = 0
    for f in files:
        try:
            if check:
                meta, cast, _, sections, _ = parse(f)
                errs, n = validate(f, meta, cast, sections)
                print(f"{f}: {n}턴 " + ("OK" if not errs else "\n  - " + "\n  - ".join(errs)))
                bad += bool(errs)
            else:
                out, n = build(f)
                print(f"{out.relative_to(ROOT.parent)}  ({n}턴)")
        except ManuscriptError as e:
            print(e)
            bad += 1
    return 1 if bad else 0


if __name__ == "__main__":
    sys.exit(main(sys.argv[1:]))
