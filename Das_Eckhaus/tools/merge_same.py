"""같은 화자의 연속 발화를 한 턴으로 합친다. 사용: python3 merge_same.py 원고.txt"""
import re, sys
p=sys.argv[1]
L=open(p,encoding='utf-8').read().rstrip('\n').split('\n'); out=[]
R=r'^([^:#]{1,30}):\s*(.+?)\s*\|\|\s*(.+)$'
for l in L:
    m=re.match(R,l)
    if m and out:
        pm=re.match(R,out[-1])
        if pm and pm.group(1)==m.group(1):
            out[-1]=f"{pm.group(1)}: {pm.group(2)} {m.group(2)} || {pm.group(3)} {m.group(3)}"; continue
    out.append(l)
open(p,'w',encoding='utf-8').write('\n'.join(out)+'\n')
