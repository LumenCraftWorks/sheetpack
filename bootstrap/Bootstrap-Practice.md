# Bootstrap practice — pseudocode

This companion to [the concept](Agent-Bootstrap-Concept.md) describes text
decoding, local tests and verification through instructions and pseudocode only.
It contains no executable source or shell commands. Follow the concept for
manifest mapping, block assembly and writing. Decoding and extraction must work
without an installed harness.

## Decode original cell text

Supply an original worksheet cell XML element and a list of original shared-string
XML elements, obtained from the workbook ZIP. XML parsing has already decoded entities.
Apply this to metadata cells too if their text could contain literal escapes.
Never supply values already modified by `openpyxl`. This avoids dependence on
private library attributes or on its version-specific string substitutions.

```text
READ ORIGINAL CELL:
    IF it contains a formula:
        REPORT an unsupported formula; do not use its cached result
    ELSE IF it references a shared string:
        READ its zero-based index
        REQUIRE the index to exist in the original shared-string table
        DECODE ORIGINAL STRING from the referenced XML element
    ELSE IF it contains an inline string:
        DECODE ORIGINAL STRING from its inline XML element
    ELSE:
        RETURN the XML-parsed scalar value as text, or empty text if absent

DECODE ORIGINAL STRING:
    IF the string element is absent:
        RETURN empty text
    COLLECT plain text and rich-text runs in document order
    IGNORE phonetic annotations and formatting
    TREAT empty text nodes as empty text
    JOIN collected text without separators
    SCAN that original text once from left to right:
        DECODE each OOXML escape into its UTF-16 code unit
        COPY other characters unchanged
        NEVER rescan characters produced by decoding
    COMBINE valid surrogate pairs into Unicode characters
    REJECT unpaired surrogates and NUL
    RETURN the decoded text without trimming or newline normalization
```

This preserves literal escapes while decoding actual escaped characters. Check it against
[the bootstrap vectors](../Bootstrap-Vectors.md). Do not clear XML elements before
their text is consumed. A shared-string index refers to the original table order.

## Start the extracted Python harness

Use Python 3.11+ and the actual extraction destination. Start the recovered package
as a module; installation is not required.

```text
SELECT the recovered harness directory as the working directory
ADD its absolute src directory to the Python module search path
LEAVE optional PowerShell integration tests disabled
START the recovered sheetpack package's help entry point as a module
DISCOVER and RUN the default Python tests in the recovered tests directory
REPORT passes, failures and skipped tests separately
```

Default discovery excludes the optional PowerShell integration tests and does not
probe for or invoke `pwsh`. Unavailable optional openpyxl tests are reported as
skipped. Report actual passes, failures and skips; do not hard-code a test count.
Do not launch `src/sheetpack/__main__.py` directly because it uses relative imports.

Python's isolated mode (`-I`) ignores
`PYTHONPATH` and does not automatically put the checkout on the import path.
An isolated bootstrap reader using only standard libraries is fine. For isolated
tests of recovered sources, explicitly arrange the recovered import path inside
that interpreter before importing the package or discovering tests.

## Verify the project workbook independently of source language

| Operation | Requires `pwsh`? |
|---|---|
| Independent bootstrap and `unpack` | No |
| Recovered package help and default Python tests | No |
| `verify`, with or without `--fix`, for Python, Java, PowerShell or other source languages | No |
| Pack a source tree including `.ps1`/`.psm1` files | Yes, for AST splitting |

Verify/Fix check the `sheetpack` workbook format, reconstructed content and
hashes. They do not parse, compile or execute source code. The included
`src/sheetpack/parse.ps1` is ordinary project content during verification;
it does not cause PowerShell to run. Perform the additional format check as follows:

```text
USE the recovered package with its src directory on the module search path
SELECT the original project workbook by its absolute path
REQUEST verification without enabling Fix
CHECK the exit status and reported errors, warnings and information
KEEP the original workbook unchanged
```

Independent extraction, the final manifest-hash comparisons and default Python
tests remain the bootstrap acceptance path. Source syntax needs a separate
language-specific check if desired. Do not use `--fix` to accept hash mismatches
caused by an incorrect extraction adapter. A regular Excel workbook without the
`sheetpack` layout is outside this verifier's format contract.
