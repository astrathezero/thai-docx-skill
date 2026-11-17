#!/usr/bin/env bash
# =============================================================================
# install.sh - Install thai-docx skill for Anthropic Claude (Claude Code)
# =============================================================================

set -e

SCRIPT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
CLAUDE_SKILLS_DIR="$HOME/.claude/skills"
AGENTS_SKILLS_DIR="$HOME/.agents/skills"

echo "=========================================================="
echo "Installing Thai DOCX Skill for Anthropic Claude..."
echo "Source: $SCRIPT_DIR"
echo "=========================================================="

# 1. Install into ~/.agents/skills if directory exists
if [ -d "$HOME/.agents" ]; then
    mkdir -p "$AGENTS_SKILLS_DIR"
    TARGET_AGENTS="$AGENTS_SKILLS_DIR/thai-docx"
    rm -rf "$TARGET_AGENTS"
    cp -R "$SCRIPT_DIR" "$TARGET_AGENTS"
    echo "✓ Installed to $TARGET_AGENTS"
fi

# 2. Install into ~/.claude/skills
mkdir -p "$CLAUDE_SKILLS_DIR"
TARGET_CLAUDE="$CLAUDE_SKILLS_DIR/thai-docx"
rm -rf "$TARGET_CLAUDE"

# Prefer symlink if installed in ~/.agents/skills, else copy
if [ -d "$AGENTS_SKILLS_DIR/thai-docx" ]; then
    ln -s "$AGENTS_SKILLS_DIR/thai-docx" "$TARGET_CLAUDE"
    echo "✓ Linked to $TARGET_CLAUDE -> $AGENTS_SKILLS_DIR/thai-docx"
else
    cp -R "$SCRIPT_DIR" "$TARGET_CLAUDE"
    echo "✓ Copied to $TARGET_CLAUDE"
fi

# 3. Install into ~/.gemini/config/skills if directory exists
GEMINI_SKILLS_DIR="$HOME/.gemini/config/skills"
if [ -d "$GEMINI_SKILLS_DIR" ]; then
    TARGET_GEMINI="$GEMINI_SKILLS_DIR/thai-docx"
    rm -rf "$TARGET_GEMINI"
    cp -R "$SCRIPT_DIR" "$TARGET_GEMINI"
    echo "✓ Installed to $TARGET_GEMINI"
fi

# 4. Check python-docx and pythainlp dependency
echo ""
echo "Checking Python environment..."
if /usr/bin/python3 -c "import docx" >/dev/null 2>&1; then
    echo "✓ python-docx is available in system Python (/usr/bin/python3)"
elif python3 -c "import docx" >/dev/null 2>&1; then
    echo "✓ python-docx is available in active python3"
else
    echo "ℹ Note: Please ensure python-docx is installed: pip install python-docx"
fi

if python3 -c "import pythainlp" >/dev/null 2>&1; then
    echo "✓ pythainlp is available for Thai word segmentation (ZWSP)"
else
    echo "ℹ Note: For best Thai word segmentation (ZWSP), install pythainlp: pip install pythainlp"
fi

echo ""
echo "=========================================================="
echo "🎉 Installation completed successfully!"
echo "Claude Code will now automatically discover 'thai-docx'!"
echo "=========================================================="
