"""원고(.txt) → 화별 대본(.docx) 변환 + 기획서 작성 상태 갱신.

원고 형식
  #nr: 1
  #part: 1부 Frühling (봄)
  #title: Ein ganz normaler Dienstag in der Ecke
  #ko: 아주 평범한 화요일
  #level: A2~B1
  #setting: 배경 설명
  #cast: Kalle | 칼레 코발스키 · 크나이페 „Zur Ecke" 주인
  #read: 낭독 포인트 한 줄
  #note: 표현 노트 한 줄 ([베를린], [속어] 태그 등)
  ## 장면 제목 | 한국어 장면 제목
  Kalle: 독일어 대사 || 한국어 해석

사용법: python3 build_episode.py 원고.txt [--no-status]
"""
import re
import sys
from pathlib import Path

from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.shared import Pt, RGBColor, Cm
from docx.oxml.ns import qn

ROOT = Path(__file__).resolve().parents[1]
NAVY, GREY, DGREY, TEXT, GOLD = (RGBColor(0x1F, 0x38, 0x64), RGBColor(0x80, 0x80, 0x80),
                                RGBColor(0x59, 0x59, 0x59), RGBColor(0x1A, 0x1A, 0x1A),
                                RGBColor(0x7F, 0x60, 0x00))
TURNS = 120


def parse(path):
    meta, cast, reads, notes, scenes = {}, [], [], [], []
    for ln, raw in enumerate(Path(path).read_text(encoding='utf-8').splitlines(), 1):
        line = raw.strip()
        if not line:
            continue
        if line.startswith('## '):
            de, _, ko = line[3:].partition('|')
            scenes.append({'de': de.strip(), 'ko': ko.strip(), 'lines': []})
        elif line.startswith('#'):
            key, _, val = line[1:].partition(':')
            key, val = key.strip(), val.strip()
            if key == 'cast':
                n, _, d = val.partition('|'); cast.append((n.strip(), d.strip()))
            elif key == 'read':
                reads.append(val)
            elif key == 'note':
                notes.append(val)
            else:
                meta[key] = val
        else:
            m = re.match(r'^([^:]{1,30}):\s*(.+?)\s*\|\|\s*(.+)$', line)
            if not m or not scenes:
                sys.exit(f'{path}:{ln}: 형식 오류 → {line[:60]}')
            scenes[-1]['lines'].append((m.group(1).strip(), m.group(2), m.group(3)))
    names = {c[0] for c in cast}
    turns = [l for s in scenes for l in s['lines']]
    bad = {l[0] for l in turns} - names
    if bad:
        sys.exit(f'{path}: 등장인물 목록에 없는 화자 {bad}')
    for s_ in scenes:
        for a, b in zip(s_['lines'], s_['lines'][1:]):
            if a[0] == b[0]:
                sys.exit(f'{path}: 같은 화자 연속 발화 → {a[0]}: {b[1][:40]}')
    if len(turns) != TURNS:
        sys.exit(f'{path}: 턴 수 {len(turns)} (필요 {TURNS})')
    for k in ('nr', 'part', 'title', 'ko', 'level', 'setting'):
        if k not in meta:
            sys.exit(f'{path}: #{k} 누락')
    return meta, cast, reads, notes, scenes


def para(doc, text='', size=11, bold=False, color=None, align=None, space_after=2):
    p = doc.add_paragraph()
    if align:
        p.alignment = align
    p.paragraph_format.space_after = Pt(space_after)
    if text:
        r = p.add_run(text); r.bold = bold; r.font.size = Pt(size)
        if color:
            r.font.color.rgb = color
    return p


def build(path):
    meta, cast, reads, notes, scenes = parse(path)
    nr, part = int(meta['nr']), meta['part']
    pno = part.split('부')[0] + '부'
    doc = Document()
    sec = doc.sections[0]
    sec.page_width, sec.page_height = Cm(21), Cm(29.7)
    for side in ('left_margin', 'right_margin'):
        setattr(sec, side, Cm(2.5))
    st = doc.styles['Normal']
    st.font.name = 'Arial'
    st.element.rPr.rFonts.set(qn('w:eastAsia'), 'Malgun Gothic')

    C = WD_ALIGN_PARAGRAPH.CENTER
    para(doc, 'DAS ECKHAUS', 15, True, GREY, C)
    para(doc, '독일어 낭독 교재 · 베를린 노이쾰른', 13, False, GREY, C)
    para(doc, f'{pno} · {nr}화', 12, True, NAVY, C)
    para(doc, meta['title'], 20, True, NAVY, C)
    para(doc, meta['ko'], 12, False, RGBColor(0x40, 0x40, 0x40), C)
    para(doc, f"대사 {TURNS}턴 · 난이도 {meta['level']} · {part}", 9, False, DGREY, C)
    para(doc, '지문 없이 대사만으로 구성 · 독일어 본문 → 낭독 포인트 → 표현 노트 → 한글 해석', 9, False, DGREY, C)
    para(doc, '등장인물', 10, True, NAVY, C)
    for n, d in cast:
        p = para(doc, align=C)
        r = p.add_run(n + '  '); r.bold = True; r.font.size = Pt(10)
        r = p.add_run(d); r.font.size = Pt(9); r.font.color.rgb = DGREY
    para(doc, '※ ' + meta['setting'], 8, False, GREY, space_after=8)

    for s in scenes:
        para(doc, s['de'], 12, True, NAVY, space_after=4)
        for who, de, _ in s['lines']:
            p = para(doc, space_after=3)
            r = p.add_run(who + ':'); r.bold = True; r.font.size = Pt(10); r.font.color.rgb = GOLD
            r = p.add_run('  ' + de); r.font.size = Pt(11); r.font.color.rgb = TEXT

    if reads:
        para(doc, '낭독 포인트', 12, True, NAVY, space_after=4)
        for x in reads:
            para(doc, '• ' + x, 10, False, TEXT)
    if notes:
        para(doc, '표현 노트', 12, True, NAVY, space_after=4)
        for x in notes:
            para(doc, '• ' + x, 10, False, TEXT)

    para(doc, '한글 해석', 12, True, NAVY, space_after=4)
    for s in scenes:
        para(doc, s['ko'] or s['de'], 10, True, DGREY, space_after=3)
        for who, _, ko in s['lines']:
            p = para(doc, space_after=2)
            r = p.add_run(who + ':'); r.bold = True; r.font.size = Pt(9); r.font.color.rgb = GOLD
            r = p.add_run('  ' + ko); r.font.size = Pt(10); r.font.color.rgb = DGREY

    safe = re.sub(r'[^A-Za-z0-9]+', '_', meta['title'].translate(str.maketrans(
        {'ä': 'ae', 'ö': 'oe', 'ü': 'ue', 'Ä': 'Ae', 'Ö': 'Oe', 'Ü': 'Ue', 'ß': 'ss'}))).strip('_')
    out = ROOT / '화별_대본' / pno / f'{pno}_{nr}화_{safe}.docx'
    out.parent.mkdir(parents=True, exist_ok=True)
    doc.save(out)
    return nr, out


def mark_done(nrs):
    import openpyxl, subprocess
    x = ROOT.parent / '기획서' / 'Das_Eckhaus_기획서.xlsx'
    wb = openpyxl.load_workbook(x)
    ws = wb['화별 개요']
    for nr in nrs:
        ws.cell(row=nr + 1, column=8).value = '초고완료'
    wb.save(x)
    recalc = Path('/root/.claude/skills/synced/35b8774b-4639-4d96-ae3c-2b86f7ed03e5_5e581076-9965-4a2d-92f6-92f092b9ed53/xlsx/scripts/recalc.py')
    if recalc.exists():
        subprocess.run([sys.executable, str(recalc), str(x), '60'], check=False, capture_output=True)


if __name__ == '__main__':
    args = [a for a in sys.argv[1:] if not a.startswith('--')]
    done = []
    for a in args:
        nr, out = build(a)
        done.append(nr)
        print(f'{nr}화 → {out.relative_to(ROOT.parent)}')
    if '--no-status' not in sys.argv:
        mark_done(done)
