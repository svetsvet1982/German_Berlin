#!/bin/bash
# 사용: tools/finish.sh 원고/1부_02화.txt 2 "커밋 제목"
set -e
cd "$(dirname "$0")/.."
python3 -I tools/build_episode.py "$1" .
python3 -I tools/mark_done.py ../기획서/Il_Cantone_이탈리아어_로마_기획서.xlsx "$2" >/dev/null
python3 /root/.claude/skills/synced/35b8774b-4639-4d96-ae3c-2b86f7ed03e5_5e581076-9965-4a2d-92f6-92f092b9ed53/xlsx/scripts/recalc.py ../기획서/Il_Cantone_이탈리아어_로마_기획서.xlsx 60 | grep -q '"total_errors": 0'
cd ..
git add -A
git commit -q -m "$3

Co-Authored-By: Claude Opus 5.5 <noreply@anthropic.com>
Claude-Session: https://claude.ai/code/session_01CehbhWYtkCUSADaKvXN9QP"
for i in 2 4 8 16; do git push -q origin claude/jolly-lovelace-k2rlai && break || sleep $i; done
git log --oneline -1
