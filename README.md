# nebo-office

Nebo plugin for document generation. Converts JSON specs into Word (.docx), Excel (.xlsx), and PowerPoint (.pptx) files. Compiled Rust binary — no runtime dependencies.

## Skills

This plugin ships three skills:

| Skill | Format | Commands |
|-------|--------|----------|
| docx  | Word   | `nebo-office docx create/unpack/validate` |
| xlsx  | Excel  | `nebo-office xlsx create/unpack/validate` |
| pptx  | PowerPoint | `nebo-office pptx create/unpack/validate` |

Each skill has its own SKILL.md in `dist/{format}/`.

## Building

```bash
cargo build --release
```

Cross-compile for all platforms:

```bash
# macOS ARM (Apple Silicon)
cargo build --release --target aarch64-apple-darwin

# macOS Intel
cargo build --release --target x86_64-apple-darwin

# Linux ARM
cross build --release --target aarch64-unknown-linux-gnu

# Linux Intel
cross build --release --target x86_64-unknown-linux-gnu

# Windows Intel
cross build --release --target x86_64-pc-windows-gnu
```

## Uploading to NeboLoop

```bash
# Get upload token via MCP binary-token call, then:
TOKEN=... ./dist/upload.sh <skill_id> <skill_dir> nebo-office
```

## License

MIT
