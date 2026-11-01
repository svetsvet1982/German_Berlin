import sys
for p in sys.argv[1:]:
    c={};cur=None
    for l in open(p,encoding='utf-8'):
        if l.startswith('#SCENE'): cur=l.strip()[7:]; c[cur]=0
        elif l.startswith('#'): cur=None
        elif cur and l.strip() and not l.startswith('//'): c[cur]+=1
    print(p, sum(c.values()), list(c.values()))
