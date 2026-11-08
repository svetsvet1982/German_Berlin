#!/bin/bash
# 사용법: commit_ep.sh <원고경로> "<커밋 제목>"
set -e
cd "$(dirname "$0")/.."
python3 tools/build_episode.py "$1"
cd ..
git add Das_Eckhaus 기획서
git commit -q -m "$2

Co-Authored-By: Claude Opus 5.5 <noreply@anthropic.com>
Claude-Session: https://claude.ai/code/session_01CehbhYCHKASrzpJVvh2XnQ"
for i in 1 2 3 4; do git push -q -u origin claude/nifty-ramanujan-m49wfh && break || sleep $((2**i)); done
git log --oneline -1
