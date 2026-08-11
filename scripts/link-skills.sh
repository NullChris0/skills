#!/usr/bin/env bash
set -euo pipefail
shopt -s nullglob

repo="$(cd "$(dirname "$0")/.." && pwd)"
dest="${AGENT_SKILLS_DIR:-$HOME/.agents/skills}"

sources=()
for skill_md in "$repo"/skills/*/SKILL.md; do
  sources+=("$(dirname "$skill_md")")
done

for src in "${sources[@]}"; do
  target="$dest/$(basename "$src")"
  if { [ -e "$target" ] || [ -L "$target" ]; } &&
    { [ ! -L "$target" ] || [[ "$(realpath -m "$target")" != "$repo/"* ]]; }; then
    echo "error: $target already exists and is not managed by $repo" >&2
    exit 1
  fi
done

mkdir -p "$dest"
for target in "$dest"/*; do
  if [ -L "$target" ] && [[ "$(realpath -m "$target")" == "$repo/skills/"* ]] &&
    [ ! -f "$repo/skills/$(basename "$target")/SKILL.md" ]; then
    rm "$target"
    echo "removed stale link $target"
  fi
done

for src in "${sources[@]}"; do
  target="$dest/$(basename "$src")"
  ln -sfn "$src" "$target"
  echo "linked $target -> $src"
done
