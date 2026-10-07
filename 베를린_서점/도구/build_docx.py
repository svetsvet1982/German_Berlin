import re,sys,glob,subprocess,os,tempfile
from docx import Document
from docx.shared import Pt, RGBColor
from docx.oxml.ns import qn
part,title=sys.argv[1],sys.argv[2]
base=os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
import tempfile
S=tempfile.mkdtemp()
files=sorted(glob.glob(f'{base}/{part}/*.md'),key=lambda p:int(re.search(r'_(\d+)화',p).group(1)))
parts=[]
for f in files:
    t=open(f,encoding='utf-8').read()
    # 줄바꿈 유지: 헤더 메타 줄(**배경:** 등)을 별도 문단으로
    t=re.sub(r'\n(\*\*(?:배경|인물|난이도|이전 줄거리|1부 줄거리 요약):\*\*)',r'\n\n\1',t)
    parts.append(t)
toc='\n'.join('- '+re.search(r'^# (.+)$',t,re.M).group(1) for t in parts)
md=f'---\ntitle: "{title}"\nsubtitle: "Berliner Buchhandlung · 독일어 대화문 + 한국어 해설"\n---\n\n**목차**\n\n{toc}\n\n'+'\n\n'.join(parts)
src=f'{S}/{part}.md'; open(src,'w',encoding='utf-8').write(md)
out=f'{S}/{part}_raw.docx'
subprocess.run(['pandoc',src,'-o',out,'-f','markdown+pipe_tables'],check=True)
d=Document(out)
def setfont(style,name='Malgun Gothic'):
    style.font.name=name
    rpr=style.element.get_or_add_rPr()
    rf=rpr.find(qn('w:rFonts'))
    if rf is None:
        from docx.oxml import OxmlElement
        rf=OxmlElement('w:rFonts'); rpr.append(rf)
    for a in ('w:ascii','w:hAnsi','w:eastAsia','w:cs'): rf.set(qn(a),name)
for st in d.styles:
    try:
        if st.type==1: setfont(st)
    except Exception: pass
for n,sz in (('Normal',10.5),('Body Text',10.5),('First Paragraph',10.5),('Compact',10)):
    try: d.styles[n].font.size=Pt(sz)
    except KeyError: pass
for st in d.styles:
    if st.name.lower()=='heading 1': st.paragraph_format.page_break_before=True
# 표 서식: 테두리 및 가는 글꼴
from docx.oxml import OxmlElement
for tbl in d.tables:
    tblPr=tbl._tbl.tblPr
    b=OxmlElement('w:tblBorders')
    for e in ('top','left','bottom','right','insideH','insideV'):
        x=OxmlElement(f'w:{e}'); x.set(qn('w:val'),'single'); x.set(qn('w:sz'),'4'); x.set(qn('w:color'),'AAAAAA'); b.append(x)
    tblPr.append(b)
    for c in tbl.rows[0].cells:
        for p in c.paragraphs:
            for r in p.runs: r.font.bold=True
final=f'{base}/워드/{title}.docx'
d.save(final); print(final)
