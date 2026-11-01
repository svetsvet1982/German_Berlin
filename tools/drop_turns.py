#!/usr/bin/env python3
"""원고에서 지정한 접두어로 시작하는 대사(풀이·대역 포함)를 삭제한다. 사용법: drop_turns.py 파일 '접두어1' '접두어2' ..."""
import sys
p=sys.argv[1]; L=open(p,encoding='utf-8').read().split('\n')
for pr in sys.argv[2:]:
    idx=[k for k,l in enumerate(L) if l.startswith(pr)]
    if len(idx)!=1: sys.exit(f"접두어 {pr!r}: {len(idx)}건 일치")
    i=idx[0]; j=i+1
    while j<len(L) and (L[j].startswith('= ') or L[j].startswith('> ')): j+=1
    del L[i:j]
open(p,'w',encoding='utf-8').write('\n'.join(L))
