---
name: officecli
description: Create, analyze, proofread, and modify Office documents (.docx, .xlsx, .pptx) using the officecli CLI tool. Use when the user wants to create, inspect, check formatting, find issues, add charts, or modify Office documents.
---

# officecli

AI-friendly CLI for .docx, .xlsx, .pptx. Single binary, no dependencies, no Office installation needed.

## Install & Update

Same command for both install and upgrade:

```bash
# macOS / Linux
curl -fsSL https://raw.githubusercontent.com/iOfficeAI/OfficeCLI/main/install.sh | bash

# Windows (PowerShell)
irm https://raw.githubusercontent.com/iOfficeAI/OfficeCLI/main/install.ps1 | iex
```

After installation, run `source ~/.zshrc` (macOS) or `source ~/.bashrc` (Linux) to make the `officecli` command available.

Verify: `officecli --version`

officecli auto-updates daily in the background.

---

## Strategy

**L1 (read) → L2 (DOM edit) → L3 (raw XML)**. Always prefer higher layers. Add `--json` for structured output.

---

## Help System (IMPORTANT)

**When unsure about property names, value formats, or command syntax, ALWAYS run help instead of guessing.** One help query is faster than guess-fail-retry loops.

**Three-layer navigation** — start from the deepest level you know:
```bash
officecli pptx set              # All settable elements and their properties
officecli pptx set shape        # Shape properties in detail
officecli pptx set shape.fill   # Specific property format and examples
```

Replace `pptx` with `docx` or `xlsx`. Commands: `view`, `get`, `query`, `set`, `add`, `raw`.

---

## Performance: Resident Mode

For multi-step workflows (3+ commands on the same file), use `open`/`close`:
```bash
officecli open report.docx       # keep in memory — fast subsequent commands
officecli set report.docx ...    # no file I/O overhead
officecli close report.docx      # save and release
```

---

## Quick Start

**PPT:**
```bash
officecli create slides.pptx
officecli add slides.pptx / --type slide --prop title="Q4 Report" --prop background=1A1A2E
officecli add slides.pptx /slide[1] --type shape --prop text="Revenue grew 25%" --prop x=2cm --prop y=5cm --prop font=Arial --prop size=24 --prop color=FFFFFF
officecli set slides.pptx /slide[1] --prop transition=fade --prop advanceTime=3000
```

**Word:**
```bash
officecli create report.docx
officecli add report.docx /body --type paragraph --prop text="Executive Summary" --prop style=Heading1
officecli add report.docx /body --type paragraph --prop text="Revenue increased by 25% year-over-year."
```

**Excel:**
```bash
officecli create data.xlsx
officecli set data.xlsx /Sheet1/A1 --prop value="Name" --prop bold=true
officecli set data.xlsx /Sheet1/B1 --prop value="Score" --prop bold=true
officecli set data.xlsx /Sheet1/A2 --prop value="Alice"
officecli set data.xlsx /Sheet1/B2 --prop value=95
```

---

## L1: Create, Read & Inspect

```bash
officecli create <file>               # Create blank .docx/.xlsx/.pptx (type from extension)
officecli view <file> <mode>          # outline | stats | issues | text | annotated
officecli get <file> <path> --depth N # Get a node and its children [--json]
officecli query <file> <selector>     # CSS-like query
officecli validate <file>             # Validate against OpenXML schema
```

### view modes

| Mode | Description | Useful flags |
|------|-------------|-------------|
| `outline` | Document structure | |
| `stats` | Statistics (pages, words, shapes) | |
| `issues` | Formatting/content/structure problems | `--type format\|content\|structure`, `--limit N` |
| `text` | Plain text extraction | `--start N --end N`, `--max-lines N` |
| `annotated` | Text with formatting annotations | |

### get

Any XML path via element localName. Use `--depth N` to expand children. Add `--json` for structured output.

```bash
officecli get report.docx '/body/p[3]' --depth 2 --json
officecli get slides.pptx '/slide[1]' --depth 1          # list all shapes on slide 1
officecli get data.xlsx '/Sheet1/B2' --json
```

Run `officecli docx get` / `officecli xlsx get` / `officecli pptx get` for all available paths.

### Stable ID Addressing

Elements with stable IDs return `@attr=value` paths instead of positional indices. These paths survive insert/delete operations — use them for multi-step workflows.

**Returned path format (output):**
```
/slide[1]/shape[@id=550950021]                    # PPT shape (cNvPr.Id)
/slide[1]/shape[@id=550950021]/paragraph[1]       # child inherits parent's @id=
/slide[1]/table[@id=1388430425]/tr[1]/tc[2]       # PPT table
/body/p[@paraId=1A2B3C4D]                         # Word paragraph
/comments/comment[@commentId=1]                    # Word comment
/footnote[@footnoteId=2]                           # Word footnote
/endnote[@endnoteId=1]                             # Word endnote
/body/sdt[@sdtId=123456]                           # Word content control
```

**All formats accepted as input** — use returned paths directly for subsequent `set`/`remove`:
```bash
officecli set slides.pptx '/slide[1]/shape[@id=550950021]' --prop bold=true
officecli set slides.pptx '/slide[1]/shape[@name=Title 1]' --prop text="New"   # @name= also works (PPT)
officecli set slides.pptx '/slide[1]/shape[2]' --prop color=red                # positional still works
```

Elements without stable IDs (slide, paragraph, run, tr/tc, row) use positional indices as fallback.

### query

CSS-like selectors: `[attr=value]`, `[attr!=value]`, `[attr~=text]`, `[attr>=value]`, `[attr<=value]`, `:contains("text")`, `:empty`, `:has(formula)`, `:no-alt`.

```bash
officecli query report.docx 'paragraph[style=Normal] > run[font!=Arial]'
officecli query slides.pptx 'shape[fill=FF0000]'
```

### validate

```bash
officecli validate report.docx    # Check for schema errors
officecli validate slides.pptx    # Must pass before delivery
```

**For large documents**, ALWAYS use `--max-lines` or `--start`/`--end` to limit output.

---

## L2: DOM Operations

### set — modify properties

```bash
officecli set <file> <path> --prop key=value [--prop ...]
```

**Any XML attribute is settable** via element path (found via `get --depth N`) — even attributes not currently present.

Run `officecli <format> set` for all settable elements. Run `officecli <format> set <element>` for detail.

**Value formats:**

| Type | Format | Examples |
|------|--------|---------|
| Colors | Hex, named, RGB, theme | `FF0000`, `red`, `rgb(255,0,0)`, `accent1`..`accent6` |
| Spacing | Unit-qualified | `12pt`, `0.5cm`, `1.5x`, `150%` |
| Dimensions | EMU or suffixed | `914400`, `2.54cm`, `1in`, `72pt`, `96px` |

### add — add elements or clone

```bash
officecli add <file> <parent> --type <type> [--prop ...]
officecli add <file> <parent> --type <type> --after <path> [--prop ...]   # insert after anchor
officecli add <file> <parent> --type <type> --before <path> [--prop ...]  # insert before anchor
officecli add <file> <parent> --type <type> --index N [--prop ...]        # insert at position (legacy)
officecli add <file> <parent> --from <path>                               # clone existing element
```

**Insert position** (`--after`, `--before`, `--index` are mutually exclusive):
- `--after "p[@paraId=1A2B3C4D]"` — insert after the anchor element (short or full path)
- `--before "/body/p[@paraId=5E6F7A8B]"` — insert before the anchor element
- `--index N` — insert at 0-based position (legacy, prefer --after/--before)
- No position flag — append to end (default)

**Element types (with aliases):**

| Format | Types |
|--------|-------|
| **pptx** | slide, shape (textbox), picture (image/img), chart, table, row (tr), connector (connection/line), group, video (audio/media), equation (formula/math), notes, paragraph (para), run, zoom (slidezoom) |
| **docx** | paragraph (para), run, table, row (tr), cell (td), image (picture/img), header, footer, section, bookmark, comment, footnote, endnote, formfield, sdt (contentcontrol), chart, equation (formula/math), field, hyperlink, style, toc, watermark, break (pagebreak/columnbreak) |
| **xlsx** | sheet, row, cell, chart, image (picture), comment, table (listobject), namedrange (definedname), pivottable (pivot), sparkline, validation (datavalidation), autofilter, shape, textbox, databar/colorscale/iconset/formulacf (conditional formatting), csv (tsv) |

**Clone:** `officecli add <file> / --from /slide[1]` — copies with all cross-part relationships.

Run `officecli <format> add` for all addable types and their properties.

### move, swap, remove

```bash
officecli move <file> <path> [--to <parent>] [--index N]
officecli swap <file> <path1> <path2>
officecli remove <file> '/body/p[4]'
```

### batch — multiple operations in one save cycle

Stops on first error by default. Use `--force` to continue past errors.

```bash
# Via stdin
echo '[
  {"command":"set","path":"/Sheet1/A1","props":{"value":"Name","bold":"true"}},
  {"command":"set","path":"/Sheet1/B1","props":{"value":"Score","bold":"true"}}
]' | officecli batch data.xlsx --json

# Via --commands (inline, no stdin needed)
officecli batch data.xlsx --commands '[{"op":"set","path":"/Sheet1/A1","props":{"value":"Done"}}]' --json

# Via --input (file)
officecli batch data.xlsx --input updates.json --force --json
```

Batch supports: `add`, `set`, `get`, `query`, `remove`, `move`, `view`, `raw`, `raw-set`, `validate`.

Batch fields: `command` (or `op`), `path`, `parent`, `type`, `from`, `to`, `index`, `props` (dict), `selector`, `mode`, `depth`, `part`, `xpath`, `action`, `xml`.

JSON output is wrapped in an envelope: `{"results": [...], "summary": {"total", "executed", "succeeded", "failed", "skipped"}}`. On error, each failed result includes the original batch item for debugging. Large outputs automatically spill to a temp file.

---

## L3: Raw XML

Use when L2 cannot express what you need. No xmlns declarations needed — prefixes auto-registered.

```bash
officecli raw <file> <part>                          # view raw XML
officecli raw-set <file> <part> --xpath "..." --action replace --xml '<w:p>...</w:p>'
officecli add-part <file> <parent>                   # create new document part (returns rId)
```

**raw-set actions:** `append`, `prepend`, `insertbefore`, `insertafter`, `replace`, `remove`, `setattr`.

Run `officecli <format> raw` for available parts per format.

---

## Common Pitfalls

| Pitfall | Correct Approach |
|---------|-----------------|
| `--name "foo"` | ❌ Use `--prop name="foo"` — all attributes go through `--prop` |
| `x=-3cm` | ❌ Negative coordinates not supported. Use `x=0cm` or `x=36cm` |
| `/shape[myname]` | ❌ Name indexing not supported. Use numeric index: `/shape[3]` |
| Guessing property names | ❌ Run `officecli <format> set <element>` to see exact names |
| Modifying an open file | ❌ Close the file in PowerPoint/WPS first |
| `\n` in shell strings | ❌ Use `\\n` for newlines in `--prop text="..."` |

---

## Specialized Skills

This skill covers the officecli CLI basics. For complex scenarios, load the dedicated skill for better results:

| Scenario | Skill | Min Version | When to Use |
|----------|-------|:-----------:|-------------|
| **Word documents** | `officecli-docx` | v1.0.23 | Create, read, edit .docx — reports, letters, memos, proposals |
| **Academic papers** | `officecli-academic-paper` | v1.0.24 | Research papers, white papers with TOC, equations, footnotes, bibliography |
| **Presentations** | `officecli-pptx` | v1.0.23 | Create, read, edit .pptx — general slide decks |
| **Pitch decks** | `officecli-pitch-deck` | v1.0.24 | Investor decks, product launches, sales decks with charts and stat callouts |
| **Morph PPT** | `morph-ppt` | v1.0.24 | Morph-animated cinematic presentations |
| **Excel** | `officecli-xlsx` | v1.0.23 | Create, read, edit .xlsx — financial models, trackers, formulas |
| **Data dashboards** | `officecli-data-dashboard` | v1.0.24 | CSV/tabular data → Excel dashboards with KPI cards, charts, sparklines |

> **How to load:** Ask your AI tool to enable the skill by name, or load the skill file from `skills/<skill-name>/SKILL.md`.

---

## Notes

- Paths are **1-based** (XPath convention): `'/body/p[3]'` = third paragraph
- `--index` is **0-based** (array convention): `--index 0` = first position
- After modifications, verify with `validate` and/or `view issues`
- **When unsure**, run `officecli <format> <command> [element[.property]]` instead of guessing
