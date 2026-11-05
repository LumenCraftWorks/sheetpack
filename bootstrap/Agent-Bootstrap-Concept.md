---
version: 1
---
# Agent Bootstrap Concept: Efficient Extraction from Excel

## Objective and prerequisites

Recover the project in a generated workbook such as `HelloWorld.xlsx` or
`Sheetpack.xlsx` into a local directory. Use `openpyxl` for
workbook access and a small text-decoding adapter where needed.
This guide contains instructions and non-executable pseudocode only.
For text-decoding pseudocode and instructions for starting the recovered harness, see
[Bootstrap practice](./Bootstrap-Practice.md). The compact
[bootstrap vectors](./Bootstrap-Vectors.md) define expected text and bytes.

Prerequisites are **Python 3.11+**, an installed `openpyxl`, read access to the
workbook and a writable parent directory. Choose a **new destination** controlled
by the agent. Do not overwrite an existing tree or modify the workbook. Excel is not required for extraction.

This is a bootstrap for complete, generated workbooks, not a general repair or
hostile-document validation tool. Omit per-block hashes, full schema auditing and
syntax analysis. Treat the input as data until extraction is verified and the
recovered code has been inspected.

## 1. Read the workbook and manifest

Load without saving. Prefer `data_only=False` so formulas remain identifiable;
report a formula in source content rather than substituting a cached value.
Require `_Manifest` and `_Config.FormatVersion` equal to `1`. Unknown versions
need an explicit migration decision.

| Manifest column | Meaning |
|---|---|
| A / B / C | Relative path / file-sheet name / `file` or `dir` |
| E / F / G | `utf-8` or `utf-8-bom` / `LF` or `CRLF` / octal file mode |
| H | SHA-256 of the complete reconstructed file bytes |
| J | EndNL: number of trailing line breaks |

Use Type C to distinguish files and empty directories; directory paths end in `/`.
For files, look up sheet B and check that H1 agrees with path A. Never derive paths
from sheet names. Require relative paths without drive prefixes, backslashes,
control characters or empty, `.` and `..` segments. Report duplicate paths and
file/directory conflicts. These small checks keep writes inside the fresh target.

## 2. Preserve decoded cell text

Do not assume that every `openpyxl` string is already the original source text.
Tests with **openpyxl 3.1.5** showed that default reading can change source files,
including `src/sheetpack/ooxml.py`: literal `_x005F_` became `_`. Rich-text
preservation did not fix this. The affected file count depends on the workbook;
always derive project file counts from the supplied manifest.

Use a tested decoding adapter for the installed library version:

- Preserve original text before lossy library escape handling. A narrowly scoped
  standard-library ZIP/XML reader can supply raw shared-string entries and their
  original cell indexes; inline strings need their original text nodes. Let the
  library handle workbook and sheet access. A full custom OOXML reader is not
  required. Record any version-bound or private-API adapter.
- Join ordinary/rich-text runs in order, decode XML entities through the XML parser,
  then decode OOXML `_xHHHH_` escapes **once**. Combine valid escaped surrogate pairs.
  Do not decode an already decoded value again.
- `_x005F_x000D_` must become literal `_x000D_`; an actual `_x000C_` escape becomes
  form-feed. If the library has lost this distinction, recover original cell text
  rather than attempting to reverse a lossy transformation.
- Preserve spaces, tabs, indentation and LF. A missing B value means empty text,
  not the word `None`. Never trim source or use CSV/display exports. Report
  unexpected CR in source cells rather than guessing its meaning.

## 3. Reconstruct logical blocks

Read rows from row 2 in physical order. Ignore only completely empty rows; an
empty B value alone does not make a row irrelevant.

| File-sheet column | Meaning |
|---|---|
| A / B / C | Signature / source content / block type |
| D / E | First row of the owning class / physical part `i/n` |
| I | Gap: leading LF count for the first top-level block; empty-line count between later blocks |

The following terms also apply to the workbook's `_Instructions` sheet:

- A **physical row/part** is one worksheet data row. A **logical block** is the
  concatenation of all its parts, without added separators.
- **LF** means one line-feed character, not a visual row in Excel. An **empty
  line** between two nonempty lines requires two LF characters in total.
- **Gap** excludes the one LF separating consecutive top-level blocks. Before
  the first block there is no such separator, so Gap counts leading LF directly.
- **EndNL** is the exact trailing LF count after all block content, before EOL
  conversion. It applies even when there are no blocks.

Group consecutive nonempty rows into complete E sequences `1/n` through `n/n`.
Concatenate B **without separators**. Never apply Gap or insert a line break between
parts. Report file and row if a sequence is incomplete or inconsistent.

If separate `method` blocks occur, attach them to the preceding `class`; use
Parent D to confirm ownership. Replace each standalone trimmed `#@METHOD@` marker
line in the class template with its corresponding complete method, in order.
Preserve method indentation; marker/method counts must agree. Do not append
methods again as top-level code. The reviewed harness has no separate method rows,
but another generated workbook may have them.

Before the first top-level block, insert Gap LF characters; before each later
block, insert Gap plus one LF characters. Missing Gap defaults to zero for the
first block and one for later blocks. Method Gap is ignored. Gap and EndNL must
be nonnegative integers.

Worked example: the first block contains `A` with Gap 0, the second contains `B`
with Gap 1, and EndNL is 0. Append `A`, then two LF characters (one separator
plus one empty line), then `B`. The result is `A\n\nB`, bytes `41 0a 0a 42`.
With Gap 0 on the second block, the result is `A\nB`. If the first block alone
has Gap 1, its result starts with `\nA`. Here `\n` denotes an actual LF, not
literal backslash and letter n. More examples are in the bootstrap vectors.

## 4. Save bytes and check completed files

Append exactly EndNL LF characters, even for a file with no blocks. For CRLF,
replace every LF with CRLF. Encode as UTF-8 and prepend the BOM only for
`utf-8-bom`. Write **bytes**, with no further newline or encoding conversion.
Create parent directories and recorded empty directories.

Read each saved file and compare its SHA-256 with manifest H. **One final hash
per file** is sufficient; omit part-hash checks. This detects reconstruction and
library-decoding errors, not just transport damage. A trusted runtime does not
establish that its reconstructed text is correct.

On a missing or mismatching hash, identify the file and stop before execution.
Keep the extraction for diagnosis; do not silently accept it, repair hashes or
claim success. Restore octal file modes where supported; report unsupported or
failed mode restoration separately.

## 5. Bootstrap workflow — pseudocode

```text
OPEN generated workbook read-only through the library
CHECK format, manifest mappings and relative destination paths
PREPARE a lossless text-decoding adapter for the installed library version
CREATE a new destination exclusively
FOR EACH manifest entry:
    IF directory: create its relative directory
    IF file:
        READ decoded rows and group complete part sequences
        EXPAND any class methods at their markers
        JOIN logical blocks using Gap, then append EndNL
        APPLY recorded EOL and UTF-8/BOM; write bytes
        CHECK saved file hash against manifest H
        RESTORE file mode where supported
REPORT file count, hash results and any limitations
HAND OVER inspected sources only after all file hashes agree
```

## 6. Handover and acceptance

Inspect the sources. For the harness, put the recovered `src` directory on the
Python module search path and start `sheetpack` **as a package module**.
Do not execute `src/sheetpack/__main__.py` directly: it uses relative imports.
A standalone `hello.py` can run as a script. Harness installation is optional.

Run the recovered tests; report passes, failures and skips separately.
The default suite validates the Python harness without looking for or invoking
`pwsh`. Do not activate optional PowerShell integration tests during bootstrap;
they require a separate explicit request through `SHEETPACK_TEST_POWERSHELL`.

The recovered CLI command `verify`, with or without `--fix`, validates the project
format, structure, reconstruction and hashes independently of source language.
Verify is an additional format check once
the harness is recovered, not a prerequisite for independent extraction. See
Bootstrap practice for execution pseudocode, dependency handling and the `-I` caveat.

Success means all manifest files and empty directories were recovered, all file
hashes agree, and the recovered package can start independently.

Qualify the approach against the actual workbook and a small set of edge cases:
literal escapes, permitted controls, multipart text, Unicode, BOM/CRLF, final
newlines, empty files and class methods. Use the linked bootstrap vectors rather
than inventing expected results. Derive the expected file count from the supplied
manifest. Hashes check content, not its author or safety.
