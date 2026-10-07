import re,sys
s=open(sys.argv[1],encoding='utf-8').read()
t=re.findall(r'^\*\*(Lena|Jonas):\*\*',s,re.M)
c=[];cur=None
for l in s.splitlines():
    if l.startswith('## Szene'): c.append(0)
    elif re.match(r'^\*\*(Lena|Jonas):\*\*',l): c[-1]+=1
print(len(t), all(t[i]!=t[i+1] for i in range(len(t)-1)), t[0], t[-1], c)
