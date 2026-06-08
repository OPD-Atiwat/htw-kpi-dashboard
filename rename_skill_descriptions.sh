#!/bin/bash
# Prefix skill descriptions with category tags
# Run this in Terminal: bash rename_skill_descriptions.sh

BASE="/var/folders/2w/32fc28yn277d9z663db65kd00000gn/T/claude-hostloop-plugins/123e33105ac53bfe/skills"

patch_desc() {
  local file="$BASE/$1/SKILL.md"
  local prefix="$2"
  if [ -f "$file" ]; then
    # Replace first occurrence of 'description:' line
    sed -i '' "s|^description: |description: ${prefix} |" "$file"
    sed -i '' "s|^  description: |  description: ${prefix} |" "$file"
    echo "✅ $1 → ${prefix}"
  else
    echo "❌ ไม่เจอ: $file"
  fi
}

# Dashboard group
patch_desc "dashboard-health-check"  "[Dashboard]"
patch_desc "howto-brain"             "[Dashboard]"
patch_desc "howto-sync-trigger"      "[Dashboard]"
patch_desc "mk13-opd-update"         "[Dashboard]"
patch_desc "tiktok-kms-mapping"      "[Dashboard]"
patch_desc "short-report"            "[Dashboard]"

# BEP group
patch_desc "bep-builder"             "[BEP]"

# ประเมิน group
patch_desc "team-kpi-commission"     "[ประเมิน]"

# เครื่องมือ
patch_desc "xlsx"                    "[เครื่องมือ]"
patch_desc "pdf"                     "[เครื่องมือ]"
patch_desc "pptx"                    "[เครื่องมือ]"
patch_desc "docx"                    "[เครื่องมือ]"
patch_desc "skill-creator"           "[System]"
patch_desc "schedule"                "[System]"

echo ""
echo "เสร็จแล้ว — restart Cowork เพื่อให้ reload skills"
