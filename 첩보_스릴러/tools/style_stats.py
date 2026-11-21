#!/usr/bin/env python3
"""문장 길이 점검 · 단문/단어 나열 대사 비율을 측정한다.

사용법: python3 style_stats.py 원고/1부/1부_01화.txt [...]
기준(문체 지침 '문장 다듬기' 참조):
  - 평균 턴 길이: B1 ≥ 13단어, B1+ ≥ 14단어, B2 ≥ 16단어
  - 문장 평균 길이 ≥ 7단어
  - 7단어 이하 턴 ≤ 12%
  - 단어 나열 턴(쉼표 2개 이상 + 동사 없이 8단어 이하) 0개
"""
import re
import statistics as st
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))
from build_docx import parse  # noqa: E402

MIN_TURN = {"B1": 13, "B1+": 14, "B2": 16}
SENT = re.compile(r"(?<=[.!?…])\s+")


def stats(path):
    meta, _, _, secs, _ = parse(path)
    words, sents, short, lists = [], [], [], []
    for s in secs:
        for spk, de, _ in s["turns"]:
            n = len(de.split())
            words.append(n)
            sents += [len(x.split()) for x in SENT.split(de) if x.strip()]
            if n <= 7:
                short.append(f"{spk}: {de}")
            if de.count(",") >= 2 and n <= 8:
                lists.append(f"{spk}: {de}")
    lvl = meta.get("level", "B1")
    res = {
        "avg_turn": round(st.mean(words), 1),
        "avg_sent": round(st.mean(sents), 1),
        "short_pct": round(100 * len(short) / len(words)),
        "lists": len(lists),
    }
    ok = (res["avg_turn"] >= MIN_TURN.get(lvl, 13) and res["avg_sent"] >= 7
          and res["short_pct"] <= 12 and res["lists"] == 0)
    return res, ok, short, lists


def main(argv):
    verbose = "-v" in argv
    bad = 0
    for f in [a for a in argv if a != "-v"]:
        res, ok, short, lists = stats(f)
        print(f"{Path(f).name}: 턴평균 {res['avg_turn']} · 문장평균 {res['avg_sent']} · "
              f"7단어↓ {res['short_pct']}% · 나열 {res['lists']} → {'OK' if ok else '기준 미달'}")
        if verbose and not ok:
            for line in lists + short:
                print("   ", line)
        bad += not ok
    return 1 if bad else 0


if __name__ == "__main__":
    sys.exit(main(sys.argv[1:]))
