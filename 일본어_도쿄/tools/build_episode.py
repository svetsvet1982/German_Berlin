"""원고(txt) → 화별 대본(docx) 변환기.

원고 형식:
  #META            key: value  (part, ep, title_ja, subtitle, info)
  #CAST            이름|설명
  #NOTE            ※ 안내문 (여러 줄 가능)
  #SCENE 제목      화자|일본어|한국어   (턴 1개 = 1줄)
  #NEXT            다음 화 예고
사용: python3 build_episode.py 원고/1부_02.txt [...]
"""
import sys, os, re
from docx import Document
from docx.shared import Pt, RGBColor, Twips
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
TURNS = 150
COLORS = {
    "鷹野": "0B5394", "春日": "A64D79", "岩城": "6D4C41", "ラン": "38761D",
    "恵": "B45F06", "静江": "7F6000", "黒崎": "990000", "堀": "666666",
    "八重": "674EA7", "久我": "1155CC", "藤堂": "45818E", "真壁": "134F5C",
    "宇佐美": "C27BA0", "氷室": "3D3D9E", "西園寺": "5B0F00", "葛西": "274E13",
    "三好": "8E7CC3", "翔太": "E69138", "レオン": "741B47",
    "江藤": "1C4587", "村井": "7F6000", "真田": "38761D", "健太": "BF9000",
    "佐々木": "A61C00", "田村": "783F04", "中西": "0C343D", "三宅": "660000",
    "葛西": "274E13", "正樹": "783F04", "若林": "A64D79", "局長": "3D3D9E",
    "局長": "434343", "西田": "A64D79", "委員長": "434343", "記者A": "666666", "記者B": "666666", "司会": "666666", "片桐": "0B5394", "森下": "38761D", "宗像": "783F04", "母": "A64D79", "フン": "783F04", "叔父": "38761D", "副総裁": "434343", "ホリス": "990000", "マーカス": "0B5394", "教官": "666666", "小池": "783F04", "拓海": "38761D", "水野": "A64D79", "審査官": "434343", "職員": "666666", "人事部長": "783F04", "権藤": "990000", "社長": "434343", "三枝": "0B5394", "一期生": "434343", "野々村": "1155CC", "課長": "434343", "委員長": "666666", "関口": "38761D", "早乙女": "A64D79", "小田": "783F04", "桐山": "0B5394", "今井": "A64D79", "母親": "B45F06", "住民": "666666", "大森": "783F04", "篠原": "990000", "調査員": "666666", "真鍋": "990000", "坂本": "783F04", "健一": "1155CC", "早瀬": "660000", "学生": "666666", "四人": "434343", "ミン": "38761D", "トゥイ": "A64D79", "アグス": "B45F06", "デウィ": "674EA7", "蓮": "0B5394", "進路指導": "666666", "看護師": "45818E", "医師": "434343", "ダニエル": "1155CC", "宮下": "783F04", "白石": "666666", "坂口": "666666", "審判長": "434343", "松岡": "666666", "榊原": "660000",
}
DEFAULT = "434343"
JA, KO = "Yu Gothic", "Malgun Gothic"


def parse(path):
    meta, cast, note, scenes, nxt = {}, [], [], [], []
    sec = None
    for raw in open(path, encoding="utf-8"):
        line = raw.rstrip("\n")
        if not line.strip() or line.startswith("//"):
            continue
        if line.startswith("#"):
            head = line[1:].split(" ", 1)
            sec = head[0]
            if sec == "SCENE":
                scenes.append((head[1].strip(), []))
            continue
        if sec == "META":
            k, v = line.split(":", 1)
            meta[k.strip()] = v.strip()
        elif sec == "CAST":
            n, d = line.split("|", 1)
            cast.append((n.strip(), d.strip()))
        elif sec == "NOTE":
            note.append(line.strip())
        elif sec == "SCENE":
            parts = line.split("|")
            if len(parts) != 3:
                raise SystemExit(f"{path}: 형식 오류: {line}")
            scenes[-1][1].append(tuple(p.strip() for p in parts))
        elif sec == "NEXT":
            nxt.append(line.strip())
    return meta, cast, " ".join(note), scenes, " ".join(nxt)


def build(path):
    meta, cast, note, scenes, nxt = parse(path)
    total = sum(len(s[1]) for s in scenes)
    if total != TURNS:
        raise SystemExit(f"{path}: 턴 수 {total} (≠{TURNS})")
    names = {n for n, _ in cast}
    for _, lines in scenes:
        for spk, ja, _ in lines:
            if spk not in names:
                raise SystemExit(f"{path}: 등장인물 목록에 없는 화자 {spk}")
            if re.search(r"[\uac00-\ud7a3]", spk + ja):
                raise SystemExit(f"{path}: 일본어 칸에 한글: {ja}")

    doc = Document()
    sec = doc.sections[0]
    sec.page_width, sec.page_height = Twips(11906), Twips(16838)
    sec.top_margin = sec.bottom_margin = Twips(1134)
    sec.left_margin = sec.right_margin = Twips(1417)
    st = doc.styles["Normal"]
    st.font.name = "Calibri"
    st.element.rPr.rFonts.set(qn("w:eastAsia"), JA)

    def para(before=None, after=0, center=False):
        p = doc.add_paragraph()
        if before is not None:
            p.paragraph_format.space_before = Twips(before)
        p.paragraph_format.space_after = Twips(after)
        if center:
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        return p

    def run(p, text, size, color=None, bold=False, ea=JA):
        r = p.add_run(text)
        r.font.size = Pt(size)
        r.bold = bold
        if color:
            r.font.color.rgb = RGBColor.from_string(color)
        r._element.get_or_add_rPr().get_or_add_rFonts().set(qn("w:eastAsia"), ea)

    part, ep = meta["part"], meta["ep"]
    run(para(2400, 200, True), "TOKYO", 15, "808080", True)
    run(para(0, 100, True), "일본어 낭독 교재", 13, "808080", ea=KO)
    run(para(600, 100, True), f"{part}부 · {ep}화", 12, "1F3864", True, ea=KO)
    run(para(0, 40, True), meta["title_ja"], 20, "1F3864", True)
    run(para(600, 40, True), meta["subtitle"], 12, "404040", ea=KO)
    run(para(0, 40, True), meta["info"], 9, "595959", ea=KO)
    run(para(0, 100, True), "지문 없이 대사만으로 구성 · 일본어 본문 + 한국어 대역", 9, "595959", ea=KO)
    run(para(500, 150, True), "등장인물", 10, "1F3864", True, ea=KO)
    for n, d in cast:
        p = para(None, 40, True)
        run(p, n + "  ", 10, None, True)
        run(p, d, 9, "595959", ea=KO)
    run(para(300, 0), note, 8, "808080", ea=KO)
    for title, lines in scenes:
        run(para(400, 240), title, 12, "1F3864", True)
        for spk, ja, ko in lines:
            p = para(160, 20)
            run(p, spk, 10, COLORS.get(spk, DEFAULT), True)
            run(p, "   " + ja, 11, "1A1A1A")
            run(para(None, 0), ko, 9, "808080", ea=KO)
    run(para(500, 100), "다음 화 예고", 12, "1F3864", True, ea=KO)
    run(para(None, 0), nxt, 9, "595959", ea=KO)

    safe = re.sub(r'[\\/:*?"<>|\s]+', "_", meta["title_ja"]).strip("_")
    outdir = os.path.join(ROOT, "화별_대본", f"{part}부")
    os.makedirs(outdir, exist_ok=True)
    out = os.path.join(outdir, f"{part}부_{ep}화_{safe}.docx")
    doc.save(out)
    print(f"ok {total}턴 → {os.path.relpath(out, ROOT)}")


if __name__ == "__main__":
    for a in sys.argv[1:]:
        build(a)
