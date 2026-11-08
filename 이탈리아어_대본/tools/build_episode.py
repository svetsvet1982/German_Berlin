"""원고(.txt) → 대본(.docx) 변환기. 턴 수(120)·화자 수를 검사한다.

원고 형식
  @part 1 / @ep 1 / @title ... / @ko ... / @level ... / @setting ... / @where ...
  @cast NAME | 설명          (여러 줄)
  @note ...                 (도입 설명)
  # 장면 제목               (장면 시작)
  NAME: 이탈리아어 || 한국어  (1턴)
  > 지문 || 한국어 지문       (턴 수에 넣지 않음)
  @points / @notes / @next  이후 줄은 해당 절의 항목 ('- '로 시작)
"""
import sys, re, os
from docx import Document
from docx.shared import Pt, RGBColor, Cm
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn

COLORS = {
 'FRANCO':'7F6000','GIULIA':'A64D79','SIMONE':'38761D','NADIA':'B45F06',
 'SETTIMIO':'6D4C41','DAVIDE':'0B5394','LOREDANA':'7F7F7F','SAMIR':'2E75B6',
 'TIZIANA':'AA3939','CHIARA':'8E44AD','RAHIM':'5B7F00','SHIRIN':'9C5700',
 'ELISA':'C0504D','MARCO':'4F6228',
}
NAVY, GREY, DGREY, TXT = '1F3864', '808080', '595959', '1A1A1A'

def parse(path):
    d = {'cast': [], 'scenes': [], 'points': [], 'notes': [], 'next': [], 'note': ''}
    sec = None
    for raw in open(path, encoding='utf-8'):
        line = raw.rstrip('\n').strip()
        if not line or line.startswith('//'):
            continue
        if line.startswith('@'):
            key, _, val = line[1:].partition(' ')
            if key in ('points', 'notes', 'next'):
                sec = key; continue
            sec = None
            if key == 'cast':
                n, _, desc = val.partition('|'); d['cast'].append((n.strip(), desc.strip()))
            else:
                d[key] = val.strip()
            continue
        if sec:
            d[sec].append(line[2:] if line.startswith('- ') else line); continue
        if line.startswith('#'):
            d['scenes'].append((line.lstrip('#').strip(), [])); continue
        if line.startswith('>'):
            it, _, ko = line[1:].partition('||')
            d['scenes'][-1][1].append(('dir', it.strip(), ko.strip())); continue
        m = re.match(r'^([A-ZÀ-Ü\.\' ]+?):\s*(.*)$', line)
        if not m:
            raise SystemExit(f'형식 오류: {line}')
        name, rest = m.group(1).strip(), m.group(2)
        if '||' not in rest:
            raise SystemExit(f'번역 누락: {line}')
        it, ko = [x.strip() for x in rest.split('||', 1)]
        if not it or not ko:
            raise SystemExit(f'빈 대사/번역: {line}')
        d['scenes'][-1][1].append(('turn', name, it, ko))
    return d

def run(p, text, size=None, bold=None, color=None, italic=None):
    r = p.add_run(text)
    r.font.name = 'Arial'
    r._element.rPr.rFonts.set(qn('w:eastAsia'), 'Malgun Gothic')
    if size: r.font.size = Pt(size)
    if bold is not None: r.bold = bold
    if italic: r.italic = True
    if color: r.font.color.rgb = RGBColor.from_string(color)
    return r

def para(doc, align=None, after=4, before=0):
    p = doc.add_paragraph()
    if align == 'c': p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.paragraph_format.space_after = Pt(after); p.paragraph_format.space_before = Pt(before)
    return p

def build(src, outdir):
    d = parse(src)
    turns = [t for _, items in d['scenes'] for t in items if t[0] == 'turn']
    speakers = sorted({t[1] for t in turns})
    castnames = [c[0] for c in d['cast']]
    errs = []
    if len(turns) != 120: errs.append(f'턴 수 {len(turns)} (120이어야 함)')
    if set(speakers) != set(castnames): errs.append(f'화자 {speakers} ≠ 등장인물 {castnames}')
    if not 2 <= len(speakers) <= 5: errs.append(f'화자 수 {len(speakers)}')
    for title, items in d['scenes']:
        prev = None
        for t in items:
            if t[0] == 'dir': prev = None; continue
            if t[1] == prev: errs.append(f'같은 화자 연속 발화: {t[1]} «{t[2][:30]}»')
            prev = t[1]
    if errs: raise SystemExit('검사 실패: ' + '; '.join(errs))

    doc = Document()
    s = doc.sections[0]
    s.page_width, s.page_height = Cm(21.0), Cm(29.7)
    s.left_margin = s.right_margin = Cm(2.5); s.top_margin = s.bottom_margin = Cm(2.0)
    run(para(doc, 'c', 0), 'ROMA', 15, True, GREY)
    run(para(doc, 'c', 6), '이탈리아어 낭독 교재 · 《Il Cantone》', 13, False, GREY)
    run(para(doc, 'c', 2), f"{d['part']}부 · {d['ep']}화", 12, True, NAVY)
    run(para(doc, 'c', 2), d['title'], 20, True, NAVY)
    run(para(doc, 'c', 6), f"{d['ko']} · {d['where']}", 12, False, '404040')
    run(para(doc, 'c', 0), f"대사 120턴 · 난이도 {d['level']} · 배경: {d['setting']}", 9, False, DGREY)
    run(para(doc, 'c', 8), '이탈리아어 본문 → 낭독 포인트 → 표현 노트 → 한국어 해석(대화 순서대로)', 9, False, DGREY)
    run(para(doc, 'c', 2), '등장인물', 10, True, NAVY)
    for n, desc in d['cast']:
        p = para(doc, 'c', 1); run(p, n + '  ', 10, True, COLORS.get(n, DGREY)); run(p, desc, 9, False, DGREY)
    if d['note']:
        run(para(doc, None, 8, 6), '※ ' + d['note'], 8, False, GREY)

    run(para(doc, None, 6, 10), '대본', 14, True, NAVY)
    for title, items in d['scenes']:
        run(para(doc, None, 4, 8), title, 12, True, NAVY)
        for t in items:
            if t[0] == 'dir':
                run(para(doc, None, 4), f'({t[1]})', 9.5, False, GREY, True)
            else:
                p = para(doc, None, 3)
                run(p, t[1] + '   ', 10, True, COLORS.get(t[1], DGREY)); run(p, t[2], 11, False, TXT)

    run(para(doc, None, 4, 14), '낭독 포인트', 12, True, NAVY)
    for x in d['points']:
        run(para(doc, None, 2), '• ' + x, 9.5, False, '333333')
    if d['notes']:
        run(para(doc, None, 4, 10), '표현 노트 — [로마말]·[속어]', 12, True, NAVY)
        for x in d['notes']:
            run(para(doc, None, 2), '• ' + x, 9.5, False, '333333')

    run(para(doc, None, 6, 14), '한국어 해석', 14, True, NAVY)
    for title, items in d['scenes']:
        run(para(doc, None, 3, 8), title, 11, True, NAVY)
        for t in items:
            if t[0] == 'dir':
                if t[2]: run(para(doc, None, 2), f'({t[2]})', 9, False, GREY, True)
            else:
                p = para(doc, None, 2)
                run(p, t[1] + '   ', 9, True, COLORS.get(t[1], DGREY)); run(p, t[3], 10, False, '404040')

    if d['next']:
        run(para(doc, None, 4, 14), '다음 화 예고', 12, True, NAVY)
        for x in d['next']:
            run(para(doc, None, 2), x, 9, False, DGREY)

    fname = f"{d['part']}부_{int(d['ep']):02d}화_{d['file']}.docx"
    sub = os.path.join(outdir, f"{d['part']}부"); os.makedirs(sub, exist_ok=True)
    out = os.path.join(sub, fname); doc.save(out)
    print(f'OK {out} · 턴 {len(turns)} · 화자 {len(speakers)}명 {speakers}')

if __name__ == '__main__':
    build(sys.argv[1], sys.argv[2])
