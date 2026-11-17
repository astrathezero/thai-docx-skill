#!/usr/bin/env bash
# =============================================================================
# install.sh - Install thai-docx skill for Google Antigravity & Anthropic Claude
# =============================================================================

set -e

SCRIPT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"

echo "=========================================================="
echo "Installing Thai DOCX Skill..."
echo "Source: $SCRIPT_DIR"
echo "=========================================================="

copy_skill() {
    local target_dir="$1"
    local desc="$2"
    mkdir -p "$target_dir"
    if command -v rsync >/dev/null 2>&1; then
        rsync -a --delete --exclude '.git' --exclude '__pycache__' --exclude '*.pyc' "$SCRIPT_DIR/" "$target_dir/"
    else
        rm -rf "$target_dir"
        mkdir -p "$target_dir"
        cp -R "$SCRIPT_DIR"/* "$target_dir/" 2>/dev/null || true
        rm -rf "$target_dir/.git" "$target_dir/__pycache__"
    fi
    echo "✓ Installed for $desc: $target_dir"
}

# 1. Google Antigravity IDE & Global Config (~/.gemini/config/skills)
copy_skill "$HOME/.gemini/config/skills/thai-docx" "Google Antigravity (IDE)"

# 2. Antigravity CLI (agy slash commands) (~/.gemini/antigravity-cli/skills)
copy_skill "$HOME/.gemini/antigravity-cli/skills/thai-docx" "Google Antigravity CLI (agy)"

# 3. Antigravity 2.0 Global (~/.agents/skills)
copy_skill "$HOME/.agents/skills/thai-docx" "Antigravity 2.0 Global (~/.agents)"

# 4. Anthropic Claude Code (~/.claude/skills)
copy_skill "$HOME/.claude/skills/thai-docx" "Anthropic Claude Code"

# 5. Check Python environment & dependencies
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
echo "Google Antigravity and Claude Code will now discover 'thai-docx'!"
echo "You can use /thai-docx in prompts and slash commands."
echo "=========================================================="
