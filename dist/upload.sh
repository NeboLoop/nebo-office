#!/bin/bash
# Usage: ./upload.sh <skill_id> <skill_dir> <binary_name>
# Requires TOKEN env var from binary-token MCP call
# Example: TOKEN="..." ./upload.sh abc-123 docx nebo-office

SKILL_ID="$1"
SKILL_DIR="$2"
BINARY="$3"
DIST="$(cd "$(dirname "$0")" && pwd)"

if [ -z "$SKILL_ID" ] || [ -z "$SKILL_DIR" ] || [ -z "$BINARY" ]; then
  echo "Usage: TOKEN=... $0 <skill_id> <skill_dir> <binary_name>"
  exit 1
fi

PLATFORMS="darwin-arm64 darwin-amd64 linux-amd64 linux-arm64 windows-amd64 windows-arm64"
SKILL_MD="$DIST/$SKILL_DIR/SKILL.md"

for PLATFORM in $PLATFORMS; do
  BINARY_PATH="$DIST/$SKILL_DIR/$PLATFORM/$BINARY"
  if [ ! -f "$BINARY_PATH" ]; then
    echo "SKIP $PLATFORM — $BINARY_PATH not found"
    continue
  fi

  echo -n "Uploading $PLATFORM... "
  RESULT=$(curl --http1.1 -s -X POST "https://neboloop.com/api/v1/developer/apps/$SKILL_ID/binaries" \
    -H "Authorization: Bearer $TOKEN" \
    -F "file=@$BINARY_PATH" \
    -F "platform=$PLATFORM" \
    -F "skill=@$SKILL_MD" 2>&1)

  if echo "$RESULT" | grep -q '"id"'; then
    echo "OK"
  else
    echo "FAIL: $RESULT"
  fi
done
