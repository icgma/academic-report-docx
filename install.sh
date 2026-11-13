#!/usr/bin/env bash
# Install `academic-report-docx` into Claude Code, Codex, dsh, ZCode, WorkBuddy, and Gemini CLI.
set -euo pipefail

SKILL_NAME="academic-report-docx"
REPO_ROOT="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
TARGET="${1:-all}"

copy_bundle() {
  local dest="$1"
  local label="$2"
  mkdir -p "$dest"
  for item in SKILL.md agents configs scripts references examples assets; do
    if [ -e "$REPO_ROOT/$item" ]; then
      cp -R "$REPO_ROOT/$item" "$dest/"
    fi
  done
  echo "[OK] Installed $SKILL_NAME -> $dest ($label)"
}

install_target() {
  case "$1" in
    claude)
      copy_bundle "$HOME/.claude/skills/$SKILL_NAME" "Claude Code"
      ;;
    codex)
      copy_bundle "$HOME/.codex/skills/$SKILL_NAME" "OpenAI Codex"
      copy_bundle "$HOME/.agents/skills/$SKILL_NAME" "Shared .agents"
      ;;
    dsh)
      copy_bundle "$HOME/.dsh/skills/$SKILL_NAME" "dsh (DeepSeek Harness)"
      ;;
    zcode)
      copy_bundle "$HOME/.zcode/skills/$SKILL_NAME" "ZCode"
      ;;
    workbuddy)
      copy_bundle "$HOME/.workbuddy/skills/$SKILL_NAME" "WorkBuddy"
      if [ -d "$HOME/.workbuddy-ai" ]; then
        copy_bundle "$HOME/.workbuddy-ai/skills/$SKILL_NAME" "WorkBuddy AI"
      fi
      ;;
    gemini)
      copy_bundle "$HOME/.gemini/config/skills/$SKILL_NAME" "Gemini / Antigravity"
      ;;
    all)
      for t in claude codex dsh zcode workbuddy gemini; do
        install_target "$t"
      done
      ;;
    *)
      echo "Usage: $0 [all|claude|codex|dsh|zcode|workbuddy|gemini]" >&2
      exit 1
      ;;
  esac
}

install_target "$TARGET"
