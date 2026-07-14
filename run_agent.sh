#!/bin/bash
set -euo pipefail

PROJECT_DIR="$HOME/requirements-agent"
LOG_FILE="$PROJECT_DIR/logs/$(date +%Y%m%d_%H%M%S).log"

cd "$PROJECT_DIR"

claude -p "inputs 폴더의 모든 CSV/XLSX를 requirements-analysis-agent 스킬 규칙에 따라
수집→정의→분석→KPI→액션→추적성까지 처리하고
outputs 폴더에 날짜별 하위 폴더를 만들어 XLSX와 Executive Summary를 저장해줘
품질 Gate를 통과하지 못한 항목은 확인 필요 목록으로 별도 정리해줘" \
  --allowedTools "Read,Write,Bash(python3:*),Bash(pip3:*)" \
  --output-format json \
  --max-turns 40 \
  > "$LOG_FILE" 2> "$PROJECT_DIR/logs/error.log"

echo "실행 완료, 로그: $LOG_FILE"
osascript -e 'display notification "요구분석 에이전트 실행 완료" with title "Requirements Agent"'
