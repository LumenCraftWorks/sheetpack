# Sheetpack

**Sheetpack is a specification for storing a text project in a single Excel workbook (`.xlsx`) and recovering it byte for byte.**

Each file becomes a worksheet, a manifest sheet records paths, encodings, line endings, file modes and SHA-256 hashes, and source text is stored in logical blocks of worksheet rows. Unpacking the workbook reproduces the original files exactly.

> This repository contains the **specification and bootstrap documentation only**. It does not include a reference implementation.

## Why

- **One file for a whole project.** Share, archive or hand over a workspace as a single workbook.
- **Language independent.** Pack, unpack, verify and fix work for any supported text in any source language, without that language's interpreter or compiler.
- **Suited for AI assistants.** Most chat-based AI assistants accept Excel documents as attachments, while ZIP archives are often handled more restrictively or rejected. A Sheetpack workbook is therefore an alternative way to hand a complete project to an assistant, and to get one back.
- **Easy for an LLM to implement.** The format is fully described in plain text, and the bootstrap guide, pseudocode and test vectors give an LLM everything needed to write an extractor or a packer and to check its own result against the SHA-256 hashes.
- **Self-recoverable.** If no Sheetpack tooling is available, a short bootstrap procedure extracts the project with only `openpyxl` and the Python standard library.
- **Lossless.** File bytes, relative paths, empty directories and recorded file modes are reconstructed exactly. Every file is verified against a SHA-256 hash.
- **No Excel needed.** Reading and writing only requires ZIP/XML access, Unicode handling and SHA-256.

## Repository layout

| Path | Content |
|---|---|
| [Sheetpack-Spec.md](Sheetpack-Spec.md) | The normative specification (workbook format version 1) |
| [bootstrap/Agent-Bootstrap-Concept.md](bootstrap/Agent-Bootstrap-Concept.md) | Step-by-step guide to extract a project from a workbook without the original tools |
| [bootstrap/Bootstrap-Practice.md](bootstrap/Bootstrap-Practice.md) | Pseudocode for text decoding, running tests and verification |
| [bootstrap/Bootstrap-Vectors.md](bootstrap/Bootstrap-Vectors.md) | Reference test vectors for decoding and byte reconstruction |

## The format at a glance

| File-system concept | Workbook representation |
|---|---|
| Project | One workbook |
| File | One file sheet |
| Relative path and file metadata | `_Manifest` and the file sheet's header |
| Empty directory | A directory row in `_Manifest` |
| Logical source or text block | One or more physical worksheet rows |
| Source text | Column B of a file sheet |

Not covered by the format: symbolic links, ownership, ACLs, timestamps and directory modes.

### Current limitation: no binary files

Format version 1 deliberately supports text files only. This is a conscious restriction of the current specification, not an oversight:

- Supported are strict UTF-8 (with or without BOM) and LF or CRLF line endings per file.
- NUL characters, unpaired surrogates, mixed line endings and lone CR are not allowed, so binary content cannot be represented.
- By default, packing rejects unsupported or binary files. With an explicit skip policy they can be left out; skipped paths and reasons are then reported, and the lossless guarantee does not apply to them.

Binary support may be considered in a later format version.

## Where to start

- **Implementing Sheetpack:** read [Sheetpack-Spec.md](Sheetpack-Spec.md). Section 1.3 defines the implementation profiles (Core, Bootstrap reader, Block editor, Syntax adapter, Executor); section 14 lists the compatibility vectors and acceptance criteria.
- **Recovering a project from a workbook:** follow [Agent-Bootstrap-Concept.md](bootstrap/Agent-Bootstrap-Concept.md) and check your decoder against [Bootstrap-Vectors.md](bootstrap/Bootstrap-Vectors.md).

## Contributing

Issues and pull requests that clarify the specification, fix errors or add test vectors are welcome.

## License

Licensed under the Apache License, Version 2.0. See [LICENSE](LICENSE) for the full text.

Copyright 2026 LumenCraftWorks Team and Contributors
