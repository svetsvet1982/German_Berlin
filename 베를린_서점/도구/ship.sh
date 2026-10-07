#!/bin/bash
# 사용: 도구/ship.sh "2부 9화 «Wohnwagen am Alex»" "상태기록에 넣을 한 줄"
cd "$(dirname "$0")/../.." || exit 1
MSG="베를린 서점 $1 독일어 대화문 및 한국어 해설 추가"
[ -n "$2" ] && echo "- $1 완료 ✔ — $2" >> 베를린_서점/진행상황.md
git add 베를린_서점
git commit -q -m "$MSG

Co-Authored-By: Claude Sonnet 5.5 <noreply@anthropic.com>
Claude-Session: https://claude.ai/code/session_01XsxMDQkWiR7f5KLPAXLaHc"
for d in 2 4 8 16 32; do
  git push -u origin claude/berliner-buchhandlung-folge-1-bjzajj 2>&1 | tail -1
  if [ -z "$(git log origin/claude/berliner-buchhandlung-folge-1-bjzajj..HEAD --oneline)" ]; then echo PUSHED; exit 0; fi
  sleep $d
done
echo "PUSH FAILED (로컬 커밋 유지)"
