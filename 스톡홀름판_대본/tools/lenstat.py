#!/usr/bin/env python3
"""턴 길이(스웨덴어 단어 수) 분포 점검: python3 tools/lenstat.py _원문/1부_01화.txt [...]"""
import re, sys, statistics as st

HEAD = {"PART", "EP", "TITLE", "SUBTITLE", "META", "CAST", "NOTE", "NEXT", "NEXTHEAD"}
BANDS = [("짧음 ≤6", 0, 6), ("중간 7–14", 7, 14), ("김 15–29", 15, 29), ("아주 김 30+", 30, 10**6)]


def words(path):
    out = []
    for m in re.finditer(r"^([A-ZÅÄÖÉ\-]+): (.*)$", open(path, encoding="utf-8").read(), re.M):
        if m.group(1) not in HEAD:
            out.append(len(m.group(2).split()))
    return out


for path in sys.argv[1:]:
    w = words(path)
    n = len(w)
    bands = "  ".join(f"{name}:{sum(lo <= x <= hi for x in w) / n:.0%}" for name, lo, hi in BANDS)
    print(f"{path}  턴={n} 평균={st.mean(w):.1f} 최대={max(w)}  {bands}")
