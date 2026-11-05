---
version: 1
---
# Bootstrap reference vectors

These data examples supplement [the bootstrap concept](Agent-Bootstrap-Concept.md).
They contain no executable source. Check an adapter against them before extracting
the actual workbook. `\n`, `\r` and `\t` below denote actual LF, CR and tab;
`_xHHHH_` spellings in the raw-input column are literal OOXML text.

## Decode original XML text exactly once

The XML parser decodes entities before OOXML escapes are processed. Inputs must
come from original XML text nodes, not already transformed library cell values.

| Raw XML text | Expected decoded value | Expected UTF-8 bytes, hex |
|---|---|---|
| `_x005F_x000D_` | Literal `_x000D_`, no CR | `5f 78 30 30 30 44 5f` |
| `_x005F_x005F_` | Literal `_x005F_` | `5f 78 30 30 35 46 5f` |
| `_x005F_x005F_x005F_x000C_` | Literal `_x005F_x000C_` | `5f 78 30 30 35 46 5f 78 30 30 30 43 5f` |
| `_x000C_` | One form-feed character | `0c` |
| `_x000D_` | One CR character | `0d` |
| `_xD83D__xDE00_` | One emoji: 😀 | `f0 9f 98 80` |
| `&amp;&lt;` | `&<` | `26 3c` |
| `_xD800_` | Error: unpaired surrogate | No output |
| `_x0000_` | Error: forbidden NUL | No output |

Decoded CR is valid as an adapter result but unexpected inside canonical source
cells. Report it at the reconstruction stage; do not confuse it with a literal
`_x000D_` in source code or with the file's CRLF setting.

For rich text, join ordinary text runs first, then decode once. For example, runs
`_x005F_` and `x000D_` together produce literal `_x000D_`. Phonetic annotations
must not become source text. Empty shared/inline strings produce empty text.

## Reconstruct exact file bytes

Unless stated otherwise: UTF-8 without BOM, LF, no methods. Content is already
decoded. Gap on each block is shown explicitly; EndNL is separate.

| Blocks and metadata | Expected file text | Expected bytes, hex |
|---|---|---|
| No blocks; EndNL 0 | Empty file | Empty byte sequence |
| No blocks; EndNL 2 | `\n\n` | `0a 0a` |
| No blocks; EndNL 0; UTF-8 BOM | BOM only | `ef bb bf` |
| `A`, Gap 0; `B`, Gap 0; EndNL 0 | `A\nB` | `41 0a 42` |
| `A`, Gap 0; `B`, Gap 1; EndNL 0 | `A\n\nB` | `41 0a 0a 42` |
| `A`, Gap 1; EndNL 0 | `\nA` | `0a 41` |
| `A`, Gap 0; `B`, Gap 1; EndNL 1 | `A\n\nB\n` | `41 0a 0a 42 0a` |
| Previous row with BOM and CRLF | BOM + `A\r\n\r\nB\r\n` | `ef bb bf 41 0d 0a 0d 0a 42 0d 0a` |
| One block from parts `A\n` and `B`; Gap 0; EndNL 0 | `A\nB`, no extra part separator | `41 0a 42` |
| One block containing space then tab; Gap 0; EndNL 0 | Space then tab, unchanged | `20 09` |

Class example: a template containing `class K {`, LF, `#@METHOD@`, LF, `}` has one
method containing `    [int] M() { return 1 }`. With Gap 0 and EndNL 1, the file
is `class K {\n    [int] M() { return 1 }\n}\n`. The four method indentation
spaces are stored content. Do not add indentation or append the method twice.

The final acceptance check is SHA-256 of each saved file against manifest H.
Matching these vectors alone does not replace that per-file check. Record the
actual workbook's file count, verified count and failures rather than copying a
file count from documentation.
