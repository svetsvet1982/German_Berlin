#!/usr/bin/env python3
"""MANDAT 원고(.txt) → 낭독 교재 docx 변환기.

원고 형식
---------
머리말 (첫 '## ' 앞):
  part: 1
  ep: 1
  title_de: Erster Tag im Westend
  title_ko: 웨스트엔드의 첫날
  scene: 월요일 아침, 프랑크푸르트 웨스트엔드 Hellweg & Cie.   (부제에 붙는 장소·시간)
  level: B1+
  setting: 월요일 오전 여덟 시, ...                              (메타줄 '배경:')
  cast: KATHARINA | 프로젝트 리더 ...                           (여러 줄 가능)
  note: ※ ...                                                    (도입 설명)
  next: 1부 2화 — «...»(...) — ...                               (다음 화 예고)
본문:
  ## 도입 — MONTAG, ACHT UHR        (장면 제목)
  KATHARINA: Guten Morgen.           (대사: 화자 대문자 + 콜론)
  > 좋은 아침.                        (바로 아래 줄: 한국어 대역)
사용:  python3 build_episode.py 원고.txt [출력디렉터리]
"""
import re, sys, os, hashlib
from docx import Document
from docx.shared import Pt, RGBColor, Emu
from docx.enum.text import WD_ALIGN_PARAGRAPH

COLORS = {
    'KATHARINA': '0B5394', 'TOBIAS': '38761D', 'SELIN': 'A64D79', 'LORENZ': '7F6000',
    'KONRAD': '6D4C41', 'CLARA': 'AA3939', 'KRÖGER': 'B45F06', 'PAUL': '134F5C',
    'UTE': '741B47', 'HENRIKE': '1155CC', 'REUTER': '666666',
}
PALETTE = ['2E75B6', '8E7CC3', 'BF9000', '45818E', 'C27BA0', '6AA84F', 'CC4125', '3D85C6', '93C47D', 'E69138']

def color_for(name):
    if name in COLORS: return COLORS[name]
    h = int(hashlib.md5(name.encode()).hexdigest(), 16)
    return PALETTE[h % len(PALETTE)]

def parse(path):
    meta, cast, sections = {}, [], []
    lines = open(path, encoding='utf-8').read().splitlines()
    i = 0
    while i < len(lines) and not lines[i].startswith('## '):
        l = lines[i].strip()
        if l and ':' in l:
            k, v = l.split(':', 1)
            k, v = k.strip(), v.strip()
            if k == 'cast':
                n, d = v.split('|', 1); cast.append((n.strip(), d.strip()))
            else:
                meta[k] = v
        i += 1
    cur = None; last = None
    for raw in lines[i:]:
        l = raw.rstrip()
        if not l.strip(): continue
        if l.startswith('## '):
            cur = {'title': l[3:].strip(), 'turns': []}; sections.append(cur); continue
        if l.startswith('>'):
            if last is None or last[2] is not None:
                raise ValueError(f'대역 위치 오류: {l}')
            last[2] = l[1:].strip(); continue
        m = re.match(r"^([A-ZÄÖÜ][A-ZÄÖÜ\-\. ]*[A-ZÄÖÜ\.]):\s+(.+)$", l)
        if not m: raise ValueError(f'해석할 수 없는 줄: {l}')
        last = [m.group(1), m.group(2).strip(), None]; cur['turns'].append(last)
    for s in sections:
        for t in s['turns']:
            if t[2] is None: raise ValueError(f'대역 누락: {t[0]}: {t[1]}')
    return meta, cast, sections

def para(doc, text='', size=None, bold=None, color=None, align=None, before=None, after=None):
    p = doc.add_paragraph()
    if align: p.alignment = align
    pf = p.paragraph_format
    if before is not None: pf.space_before = Emu(before)
    if after is not None: pf.space_after = Emu(after)
    if text:
        r = p.add_run(text)
        if size: r.font.size = Pt(size)
        if bold is not None: r.bold = bold
        if color: r.font.color.rgb = RGBColor.from_string(color)
    return p

def fname_part(s):
    s = s.replace('ä','ae').replace('ö','oe').replace('ü','ue').replace('Ä','Ae').replace('Ö','Oe').replace('Ü','Ue').replace('ß','ss')
    s = re.sub(r"[^A-Za-z0-9 \-]", '', s)
    return re.sub(r"\s+", '_', s.strip())

def build(path, outdir=None):
    meta, cast, sections = parse(path)
    nturns = sum(len(s['turns']) for s in sections)
    part, ep = meta['part'], meta['ep']
    doc = Document()
    sec = doc.sections[0]
    sec.page_width, sec.page_height = Emu(7560310), Emu(10692130)
    sec.left_margin = sec.right_margin = Emu(899795)
    sec.top_margin = sec.bottom_margin = Emu(720090)
    C = WD_ALIGN_PARAGRAPH.CENTER
    para(doc, 'MANDAT', 15, True, '808080', C, 1524000, 127000)
    para(doc, '독일어 낭독 교재 · Frankfurt · Berlin', 13, False, '808080', C, 0, 63500)
    para(doc, f'{part}부 · {ep}화', 12, True, '1F3864', C, 381000, 63500)
    para(doc, meta['title_de'], 20, True, '1F3864', C, 0, 25400)
    para(doc, f"{meta['title_ko']} · {meta['scene']}", 12, False, '404040', C, 381000, 25400)
    para(doc, f"대사 {nturns}턴 · 난이도 {meta['level']} · 배경: {meta['setting']}", 9, False, '595959', C, 0, 25400)
    para(doc, '지문 없이 대사만으로 구성 · 독일어 본문 + 한국어 대역', 9, False, '595959', C, 0, 63500)
    para(doc, '등장인물', 10, True, '1F3864', C, 317500, 95250)
    for n, d in cast:
        p = para(doc, align=C, after=25400)
        r = p.add_run(n + '  '); r.bold = True; r.font.size = Pt(10)
        r = p.add_run(d); r.font.size = Pt(9); r.font.color.rgb = RGBColor.from_string('595959')
    para(doc, meta.get('note', ''), 8, None, '808080', None, 0, 0)
    doc.add_paragraph()
    for s in sections:
        para(doc, s['title'], 12, True, '1F3864', None, 254000, 152400)
        for n, de, ko in s['turns']:
            p = para(doc, before=101600, after=12700)
            r = p.add_run(n); r.bold = True; r.font.size = Pt(10); r.font.color.rgb = RGBColor.from_string(color_for(n))
            r = p.add_run('   ' + de); r.font.size = Pt(11); r.font.color.rgb = RGBColor.from_string('1A1A1A')
            para(doc, ko, 9, None, '808080', None, None, 0)
    if meta.get('next'):
        para(doc, '다음 화 예고', 12, True, '1F3864', None, 317500, 63500)
        para(doc, meta['next'], 9, None, '595959')
    outdir = outdir or os.path.join(os.path.dirname(os.path.abspath(__file__)), '..', '화별_대본', f'{part}부')
    os.makedirs(outdir, exist_ok=True)
    out = os.path.join(outdir, f"{part}부_{ep}화_{fname_part(meta['title_de'])}.docx")
    doc.core_properties.title = f"MANDAT {part}부 {ep}화 – {meta['title_de']}"
    doc.save(out)
    return out, nturns

if __name__ == '__main__':
    out, n = build(sys.argv[1], sys.argv[2] if len(sys.argv) > 2 else None)
    print(f'{os.path.normpath(out)}  ({n}턴)')
    if n != 120: print(f'경고: 턴 수가 120이 아님 ({n})'); sys.exit(2)
