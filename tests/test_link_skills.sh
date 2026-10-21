#!/usr/bin/env bash
set -euo pipefail

repo="$(cd "$(dirname "$0")/.." && pwd)"
sandbox="$(mktemp -d)"
trap 'rm -rf "$sandbox"' EXIT
dest="$sandbox/skills"

AGENT_SKILLS_DIR="$dest" "$repo/scripts/link-skills.sh"

test -L "$dest/nesma-count"
test "$(readlink "$dest/nesma-count")" = "$repo/skills/nesma-count"

collision_dest="$sandbox/collision"
mkdir -p "$collision_dest/nesma-count"
touch "$collision_dest/nesma-count/keep"
if AGENT_SKILLS_DIR="$collision_dest" "$repo/scripts/link-skills.sh"; then
  echo "expected a conflicting directory to stop the sync" >&2
  exit 1
fi
test -f "$collision_dest/nesma-count/keep"

ln -s "$repo/skills/removed-skill" "$dest/removed-skill"
AGENT_SKILLS_DIR="$dest" "$repo/scripts/link-skills.sh"
test ! -L "$dest/removed-skill"
