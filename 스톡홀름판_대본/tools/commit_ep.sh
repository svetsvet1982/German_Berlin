#!/bin/bash
# 사용: commit_ep.sh "커밋 제목"
set -e
cd "$(dirname "$0")/../.."
git add 스톡홀름판_대본
git commit -q -m "$1

Co-Authored-By: Claude Opus 5.5 <noreply@anthropic.com>
Claude-Session: https://claude.ai/code/session_01CehbhX1MCZk9yaJtWxcgPZ"
for i in 1 2 3 4; do git push -q origin claude/hopeful-thompson-pxwh3s && break || sleep $((2**i)); done
echo "pushed: $1"
