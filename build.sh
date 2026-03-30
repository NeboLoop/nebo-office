#!/usr/bin/env bash
set -euo pipefail

# Cross-compiles nebo-office for all platforms and updates plugin.json with SHA256 hashes.
# Usage: ./build.sh
# Requires: cargo, cross (for Linux/Windows targets)

BINARY="nebo-office"
DIST="dist/plugin"

declare -A TARGET_MAP=(
  ["macos-arm64"]="aarch64-apple-darwin"
  ["macos-amd64"]="x86_64-apple-darwin"
  ["linux-arm64"]="aarch64-unknown-linux-gnu"
  ["linux-amd64"]="x86_64-unknown-linux-gnu"
  ["windows-amd64"]="x86_64-pc-windows-gnu"
)

# Read version from Cargo.toml workspace
VERSION=$(grep '^version' Cargo.toml | head -1 | sed 's/.*"\(.*\)".*/\1/')
echo "Building $BINARY v$VERSION for all platforms..."

mkdir -p "$DIST"

for PLATFORM in "${!TARGET_MAP[@]}"; do
  TARGET="${TARGET_MAP[$PLATFORM]}"
  OUT_DIR="$DIST/$PLATFORM"
  mkdir -p "$OUT_DIR"

  echo -n "  $PLATFORM ($TARGET)... "

  # Use cargo for native macOS, cross for Linux/Windows
  if [[ "$TARGET" == *"apple"* ]]; then
    cargo build --release --target "$TARGET" 2>/dev/null
  else
    cross build --release --target "$TARGET" 2>/dev/null
  fi

  # Copy binary
  if [[ "$PLATFORM" == "windows"* ]]; then
    BIN_NAME="${BINARY}.exe"
    cp "target/$TARGET/release/${BINARY}.exe" "$OUT_DIR/$BIN_NAME"
  else
    BIN_NAME="$BINARY"
    cp "target/$TARGET/release/$BINARY" "$OUT_DIR/$BIN_NAME"
    chmod 755 "$OUT_DIR/$BIN_NAME"
  fi

  BINARY_PATH="$OUT_DIR/$BIN_NAME"

  # Compute SHA256
  SHA256=$(shasum -a 256 "$BINARY_PATH" | awk '{print $1}')
  SIZE=$(stat -f%z "$BINARY_PATH" 2>/dev/null || stat -c%s "$BINARY_PATH")

  echo "sha256=$SHA256 size=$SIZE"

  # Update plugin.json
  jq --arg p "$PLATFORM" \
     --arg sha "$SHA256" \
     --arg size "$SIZE" \
     --arg bn "$BIN_NAME" \
     --arg ver "$VERSION" \
     '.version = $ver | .platforms[$p].sha256 = $sha | .platforms[$p].size = ($size | tonumber) | .platforms[$p].binaryName = $bn' \
     plugin.json > plugin.json.tmp && mv plugin.json.tmp plugin.json
done

echo ""
echo "Done. plugin.json updated with version $VERSION."
echo "Binaries in $DIST/:"
ls -lR "$DIST"/
