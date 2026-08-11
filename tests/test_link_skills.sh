#!/usr/bin/env bash
set -euo pipefail

repo="$(cd "$(dirname "$0")/.." && pwd)"
sandbox="$(mktemp -d)"
trap 'rm -rf "$sandbox"' EXIT
dest="$sandbox/skills"

AGENT_SKILLS_DIR="$dest" "$repo/scripts/link-skills.sh"

test -L "$dest/estimate-with-nesma"
test "$(readlink "$dest/estimate-with-nesma")" = "$repo/skills/estimate-with-nesma"

collision_dest="$sandbox/collision"
mkdir -p "$collision_dest/estimate-with-nesma"
touch "$collision_dest/estimate-with-nesma/keep"
if AGENT_SKILLS_DIR="$collision_dest" "$repo/scripts/link-skills.sh"; then
  echo "expected a conflicting directory to stop the sync" >&2
  exit 1
fi
test -f "$collision_dest/estimate-with-nesma/keep"

ln -s "$repo/skills/removed-skill" "$dest/removed-skill"
AGENT_SKILLS_DIR="$dest" "$repo/scripts/link-skills.sh"
test ! -L "$dest/removed-skill"
