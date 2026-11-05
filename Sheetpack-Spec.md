---
version: 1
---
# Sheetpack specification: a text project in one Excel workbook

**Workbook format version: 1.** This document is a standalone specification of the
workbook representation, reconstruction algorithms, operations and acceptance
criteria. It does not depend on a particular repository, programming language,
runtime, library or command-line interface.

## 1. Scope, requirements and implementation profiles

### 1.1 Goal

Represent a directory of supported text files in a single `.xlsx` workbook:

| File-system concept | Workbook representation |
|---|---|
| Project | One workbook |
| File | One file sheet |
| Relative path and file metadata | `_Manifest` and the file sheet's H1 |
| Empty directory | A directory row in `_Manifest` |
| Logical source or text block | One or more physical worksheet rows |
| Source text | Column B of a file sheet |

The principal guarantee is **lossless reconstruction of included file bytes**,
relative paths and empty directories, plus recorded file modes where supported.
The guarantee excludes filtered or intentionally skipped files, unsupported
encodings, symbolic links, ownership, ACLs, timestamps and directory modes.

File-byte identity, logical-block identity and workbook ZIP-byte identity are
separate properties. Repacking identical files need not reproduce identical ZIP
bytes, styles or block boundaries. Arbitrary spreadsheet features are outside
the source-content guarantee.

This is a project format, not a validator for every ordinary Excel document.
Binary files are not represented by format 1.

### 1.2 Normative terms

- **MUST / MUST NOT:** required for the stated profile or operation.
- **SHOULD:** recommended; a deviation must have a documented reason.
- **MAY:** optional behavior whose availability must be discoverable.
- **Physical part:** one nonempty data row in a file sheet.
- **Logical block:** all physical parts of one block, concatenated in order.
- **Top-level block:** a block emitted into the file directly, including an expanded
  class template. A separate method is not a top-level block.
- **LF / CR:** Unicode U+000A / U+000D. These are characters, not visual Excel rows.
- **UTF-16 unit:** one 16-bit code unit; a non-BMP character uses two units.
- **Source language:** descriptive metadata about file contents. It does not
  select a runtime or authorize execution during reconstruction or verification.

Unqualified field names and column positions below are normative English format
identifiers. User-facing messages and implementation identifiers may be localized.

### 1.3 Profiles and priorities

| Profile | Required capability |
|---|---|
| Core | Read/write the format; Pack, Unpack, Verify and explicit Fix |
| Bootstrap reader | Extract a complete generated workbook independently of the tools stored inside it; check final file hashes |
| Block editor | Core plus List/Get/Set/Add/Remove logical blocks and hash preconditions |
| Syntax adapter | Optional source-language analysis for splitting or validating edits |
| Executor | Optional execution of an explicitly selected reconstructed project entry point |

A Core implementation MUST process supported text in any source language without
requiring that language's interpreter, compiler or parser. Text splitting is the
always-available baseline. A Core reader MUST understand all block types and
method-template reconstruction even if its writer only creates `text` blocks.

Syntax adapters and execution are separate capabilities. Their dependencies MUST
NOT become dependencies of Verify, Fix, Unpack or bootstrap extraction. An
implementation may provide operations as library functions, commands, services
or agent tools; the names in this document describe behavior, not required syntax.

## 2. Runtime primitives and XLSX access

No Excel installation is required. An implementation needs ZIP/package access,
namespace-aware XML reading/writing, Unicode and UTF-8 handling, SHA-256, and local
filesystem operations. A spreadsheet library may supply these facilities, but
its decoded values MUST be checked for lossless source-text handling.

### 2.1 Package structure

A writer SHOULD use the following conventional part names. Readers MUST resolve
relationships rather than infer worksheet part names from sheet order or names.

| Part | Purpose |
|---|---|
| `[Content_Types].xml` | Package content types |
| `_rels/.rels` | Package relationship to the workbook |
| `xl/workbook.xml` | Sheet names and workbook relationship IDs |
| `xl/_rels/workbook.xml.rels` | Worksheet, string-table and style relationships |
| `xl/worksheets/sheetN.xml` | Worksheet cells and presentation |
| `xl/sharedStrings.xml` | Optional shared-string table |
| `xl/styles.xml` | Optional styles, including text number format |

Required namespace identifiers for conventional XLSX parts are:

| Use | Namespace |
|---|---|
| Spreadsheet elements | `http://schemas.openxmlformats.org/spreadsheetml/2006/main` |
| Document relationship attributes | `http://schemas.openxmlformats.org/officeDocument/2006/relationships` |
| Package relationship elements | `http://schemas.openxmlformats.org/package/2006/relationships` |
| Content types | `http://schemas.openxmlformats.org/package/2006/content-types` |
| XML whitespace attribute | `http://www.w3.org/XML/1998/namespace` |

Document relationship types use the document-relationship namespace followed by
`/officeDocument`, `/worksheet`, `/sharedStrings` or `/styles` as applicable.
Relative targets are resolved against the containing part's directory. Absolute
package targets start at the ZIP root. A reader MUST NOT follow external targets
for project sheets or shared strings. Reject duplicate part names, ambiguous
relationships and targets escaping the package.

Conventional content types are `application/xml` for generic XML and
`application/vnd.openxmlformats-package.relationships+xml` for relationship parts.
Workbook, worksheet, shared-string and style overrides respectively use:

- `application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml`
- `application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml`
- `application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml`
- `application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml`

An implementation using an XLSX library need not assemble these parts manually,
but MUST produce a valid package and preserve the format's decoded values.
DTD/entity declarations MUST NOT be used to expand external or executable input.
Package reading MUST NOT execute workbook formulas, macros or project code.

The relevant XML element paths and attributes, using the namespaces above, are:

| Part | Element path | Required interpretation |
|---|---|---|
| Workbook | `workbook/sheets/sheet` | `name` is the visible name; `sheetId` is a unique numeric ID; document-namespace `id` selects its relationship |
| Relationships | `Relationships/Relationship` | `Id`, `Type` and `Target` map a resource; `TargetMode="External"` must not be followed for project data |
| Worksheet | `worksheet/sheetData/row` | `r` is the physical row number |
| Worksheet | `worksheet/sheetData/row/c` | `r` is the A1 cell address; `t` selects its storage type |
| Shared strings | `sst/si` | The ordered `si` elements define zero-based indexes |
| Content types | `Types/Default` | `Extension` and `ContentType` provide an extension mapping |
| Content types | `Types/Override` | `PartName` and `ContentType` identify a specific part |

For a text cell, `t="s"` uses a `v` child containing its decimal shared-string
index. `t="inlineStr"` uses an `is` child containing `t` or `r/t` text. A formula
is an `f` child, irrespective of any cached `v`. A1 addresses combine a column
label (`A` through `Z`, then `AA`, etc.) and a positive decimal row. A cell's row
component must match its enclosing row. Readers must not depend on XML prefixes,
which are arbitrary aliases for the namespace identifiers.

### 2.2 Reading cells

A reader MUST support shared strings and inline strings, including rich text:

1. Locate cells by explicit row/column addresses, not by the number of returned
   nonempty values. Rows and columns are one-based; shared-string indexes are
   zero-based.
2. For a shared-string cell, read its index from `v` and select the original
   `si` entry from the string table. Negative or out-of-range indexes are errors.
3. For an inline string, read its `is` element.
4. Within either string form, concatenate ordinary `t` text and rich-text `r/t`
   runs in document order. Ignore phonetic annotations and formatting.
5. XML parsing decodes entities. Decode OOXML escapes exactly once after joining
   text runs, as defined in section 5.
6. Empty string elements or absent B values mean empty text, not a textual null
   value. Scalar metadata values MAY be read from their original XML value as
   text; do not apply date or locale-dependent conversions.

Writers MUST store source and metadata as text, never as formulas. A value
starting with `=`, `+`, `-` or `@` remains literal text. Readers MUST reject
formulas in file sheets, including extension columns, and in authoritative
manifest/configuration cells. Cached formula results are not source text.

Some libraries transform OOXML escapes while reading. Applying another unescape
operation to an already transformed value is not safe. If the library loses
information, obtain original text nodes and string indexes through a small
adapter or a direct package reader. Qualify that path with section 14's vectors.

### 2.3 Workbook and presentation limits

| Limit | Requirement |
|---|---|
| Cell text | At most 32,767 UTF-16 units in every emitted cell |
| Sheet name | At most 31 UTF-16 units; naming rules in section 3.4 |
| Worksheet size | At most 1,048,576 rows and 16,384 columns |
| Source-cell line endings | LF |
| ChunkSize | Integer 1000 through 32000; default 30000 |

Never truncate source to fit a limit. Split B content into parts; reject other
unrepresentable fields with a clear diagnostic. Measure decoded Unicode text,
not the longer XML entity/escape spellings. Do not split surrogate pairs.

Writers SHOULD freeze the header, use an autofilter, a monospace source font,
text number format, wrapped source cells and top alignment. File sheets SHOULD
be ordered by path; tab colors MAY distinguish directories. Presentation is not
source content. Row-height or viewer clipping does not establish cell completeness.

## 3. Workbook schemas and paths

### 3.1 System sheets and configuration

Writers MUST emit `_Manifest`, `_Config` and `_Instructions` with that spelling.
Readers MUST recognize these names case-insensitively. Other underscore-prefixed
sheets are auxiliary and MUST NOT be treated as project files. All other sheets
are file sheets and MUST have a valid path in H1.

`_Config` has `Key` and `Value` in A1:B1, followed by one pair per row:

| Key | Meaning / writer default |
|---|---|
| FormatVersion | Decimal `1` |
| ProjectName | Explicit name, otherwise the source directory basename; on recovery, the workbook basename |
| ChunkSize | Decimal integer 1000..32000; default `30000` |
| Language | Descriptive project metadata; default `text` |

Keys MUST be unique; canonical key spelling is significant. Unknown key/value
pairs MUST survive Fix and full export. `Language` MUST NOT override per-file
metadata or cause verification to load a parser.

An unsupported FormatVersion is an error before any operation mutates an input
workbook. Never silently downgrade it during Fix. Migration is a separate,
explicitly requested operation. Missing metadata is addressed only by the
recovery rules in section 9; it is not evidence of a different format version.

`_Instructions` is a one-column guide: a title in A1 and conventions in following
rows. Its exact prose is not normative. It SHOULD explain text-only cells, full
read-back, parts versus blocks, exact Gap/EndNL semantics, class markers, hash
maintenance and the separation between format verification and syntax analysis.
Its absence does not prevent reconstruction when the authoritative data exist.

### 3.2 Manifest

`_Manifest` has these eleven headers in row 1. Empty fields are empty cells.

| Column | Header | Meaning |
|---|---|---|
| A | Path | Unique relative file path, or directory path ending in `/` |
| B | Sheet | Actual file-sheet name; empty for directories |
| C | Type | `file` or `dir` |
| D | Language | Descriptive file-language identifier; default `text` |
| E | Encoding | `utf-8` or `utf-8-bom`; default `utf-8` |
| F | EOL | `LF` or `CRLF`; default `LF` |
| G | Mode | Three or four octal digits without a prefix; default `644` |
| H | SHA-256 | Lowercase 64-digit hexadecimal hash of reconstructed file bytes |
| I | Parts | Physical block-row count, including methods and continuations; `0` for a file with no blocks |
| J | EndNL | Nonnegative integer counting trailing LF before EOL conversion; default `1` for a new file sheet |
| K | Flags | Empty or `hard-split` |

A directory row carries only Path and Type. Empty directories require such rows;
nonempty directories are implied by file paths. Writers SHOULD omit redundant
nonempty-directory entries; readers MAY accept them if the path tree is consistent.

Each file in a complete generated workbook MUST have exactly one file sheet and
one manifest row. H1 is authoritative for the path. Encoding, EOL, Mode and EndNL
come from the manifest row matching that exact path. Sheet B is a derived lookup.
Do not recover a path from a sanitized sheet name.

Changing H1 to another path creates a new path for reconciliation: metadata from
the old path MUST NOT silently transfer. An explicit rename operation MAY carry
metadata, but must declare that behavior. New file sheets without metadata use
the listed defaults only under explicit recovery/Fix, with a diagnostic.

Language identifiers do not constrain content or verification. Optional examples
of extension classification are `.java` to `java`, `.py` to `python`, `.ps1` to
`powershell`, and `.md` to `markdown`. Implementations MAY supply such mappings,
but unknown extensions and unknown language identifiers MUST remain readable as
text. No named source language is required by this specification.

### 3.3 Project paths and feasible trees

Project paths MUST:

- Be nonempty, relative and stored using `/` separators.
- Contain no backslash, U+0000 through U+001F control, absolute prefix or drive
  prefix such as `C:`; have no surrounding whitespace.
- Contain no empty, `.` or `..` segment.
- End in exactly one `/` for a directory row and no `/` for a file.

Paths are exact Unicode identifiers; do not normalize their spelling or fold
case while reconstructing files. Exact duplicates are errors. A file cannot also
be a directory or an ancestor of another file. Check all entries as a tree.

Case-only differences, reserved host names/characters, trailing segment spaces
or dots and long paths MUST receive portability diagnostics. A warning does not
make a path valid on every host. Unpack MUST fail before writing if the chosen
filesystem cannot represent the requested tree without collisions or renaming.
A relative-path length above 240 characters SHOULD trigger an advisory; it is not
a universal host limit or a guarantee about the full destination path.

### 3.4 Sheet names and allocation

File-sheet names MUST be nonempty, unique case-insensitively, at most 31 UTF-16
units, without `\ / ? * [ ] :` or control characters. They MUST NOT start or end
with an apostrophe, start with `_`, or equal `History` case-insensitively. System
and auxiliary sheet names reserve their names before file names are allocated.
Use locale-independent Unicode case folding for name/selector comparisons;
project-path lookup remains exact. A writer must also respect its XLSX library's
sheet-name restrictions.

A deterministic allocation algorithm is:

1. Process paths in ordinal UTF-16-unit order. Reserve valid existing names first
   during Fix so new allocations do not steal them.
2. Replace `/` with `.`, then other invalid sheet-name characters with `_`.
   Trim boundary apostrophes; replace an initial `_` with `-`; change `History`
   to `History~`. Use `file` if the candidate becomes empty.
3. If the candidate fits and is unused, keep it.
4. Otherwise take at most its first 26 UTF-16 units, without breaking a surrogate
   pair, and append `~` plus the first four lowercase hexadecimal digits of
   SHA-256 over the exact path's UTF-8 bytes without BOM.
5. On another collision, append the same hash plus a decimal counter starting at
   `2`, preceded by `~`. Keep at most 24 units of the candidate prefix, reducing
   it further when necessary to keep the complete name within 31 units.
6. Continue until the name is unused. Revalidate the complete result.

Allocation is technical presentation, not file identity. Other valid unique names
are allowed. Fix MUST retain valid, unambiguous existing names where possible.

## 4. File sheets, logical blocks and methods

### 4.1 Fixed header and data columns

The header row is:

| Column | Row 1 | Data rows |
|---|---|---|
| A | Signature | Navigation signature, including a generated part suffix when needed |
| B | Content | Source text in LF form |
| C | Type | Logical block type |
| D | Parent | First physical row of the owning class for methods; otherwise empty |
| E | Part | `i/n`, including `1/1` for an unsplit block |
| F | Hash | SHA-256 over this part's decoded B text encoded as UTF-8 without BOM |
| G | Path | Unused |
| H | Actual relative file path | Unused |
| I | Gap | Leading gap on the first part of a non-method block; otherwise empty |

Row numbers are physical worksheet addresses. Duplicate addresses are errors.
Ignore only completely empty data rows, including extension cells when deciding
emptiness. An empty B value alone MUST NOT cause a row to be skipped. A row with
only an annotation but no required block fields is not a valid block row.

The fixed types are `text`, `toplevel`, `function`, `filter`, `class`, `enum`,
`interface` and `method`. These are representation roles, not grammar validators.
All except `method` are emitted as top-level blocks. Only `class` permits separate
method rows and template expansion. Unknown types require an explicit format
extension; do not silently infer their reconstruction semantics.

Signatures label blocks; B contains the complete original text, including any
source declaration, comments and indentation. A signature is not executable code
and need not equal a source-language declaration. Signatures generated by tools
SHOULD be limited to 200 UTF-16 units and have whitespace runs collapsed.
Duplicate generated signatures SHOULD receive ` #2`, ` #3`, etc., with truncation
adjusted to retain the limit. Existing duplicate labels are diagnostic, not proof
of missing content; selectors must report ambiguity.

### 4.2 Multipart sequences

A complete block is a sequence of nonempty rows with Part values `1/n` through
`n/n`, where `1 <= i <= n`. Rows may have unused physical addresses between them,
but MUST be consecutive in the logical sequence of nonempty rows.

For one part, A contains the base signature. For multiple parts, A contains the
base followed by ` [Part i/n]`. The Type and base signature MUST agree across the
sequence. Remove only a terminal display suffix matching that row's Part value
when assembling the logical signature; never remove text from B.

Concatenate B values **without inserting any separator**. Gap and block-level
metadata occur on the first part only. Parent on every method part points to the
same owning class's first physical row. A broken, out-of-order or incomplete
sequence is an error; missing source text cannot be reconstructed by guessing.

### 4.3 Class templates and separate methods

A class block may contain standalone lines whose trimmed text is `#@METHOD@`.
This is a reserved representation marker, regardless of source language. A source
line with that exact standalone spelling cannot simultaneously be ordinary class
content; a writer must reject that ambiguous use rather than change its bytes.

The class's parts are followed by all its method blocks, before the next non-method
block. Adjacency defines ownership; Parent redundantly identifies the same class.
A method without a preceding class is an error. A Parent mismatch is invalid
derived metadata, not permission to attach the method to a different class.

The first marker line is replaced by the first method's complete content, the
second by the second, and so on. No indentation or separator is added. All other
class-template lines are retained. The number of markers MUST equal the number
of method blocks for normal reconstruction. Methods are not appended again.

A template is intermediate data and need not be valid syntax in a source language.
Core tools MUST NOT require a parser to expand it. Constructors, properties and
unextracted members simply remain in the template.

## 5. Text serialization and character handling

### 5.1 Supported source bytes

Format 1 supports strict UTF-8, either without BOM or with one leading byte
sequence `EF BB BF`. It supports per-file LF or CRLF. Pack MUST reject mixed EOL
and lone CR in imported source, unless the file is explicitly skipped as
unsupported. NUL and unpaired surrogates are forbidden. Other permitted Unicode
characters and escaped XML-incompatible controls must survive exactly.

This is an encoding limitation, not a source-language limitation. An unsupported
text encoding MUST NOT be silently converted while claiming byte identity.

### 5.2 Writing and reading OOXML escapes

A writer MUST preserve literal escape spellings and allowed control characters:

1. Inspect the original text. For every underscore that begins a literal
   `_xHHHH_` spelling, replace that underscore with `_x005F_`, keeping the following
   characters. Check each original underscore, including overlapping spellings;
   a consuming replacement must not miss a shared boundary underscore.
2. Encode XML-incompatible controls and noncharacters U+FFFE/U+FFFF with OOXML
   `_xHHHH_` escapes. Do not reprocess output created in step 1.
3. Apply ordinary XML entity escaping. Preserve text-node boundary whitespace
   using `xml:space="preserve"`. Literal metadata CR must also survive XML
   newline normalization, for example through an OOXML escape or character reference.
4. Emit text cells, using shared or inline strings as appropriate.

Reading reverses those layers in this order: XML parsing, joining ordinary text
runs, **one** OOXML-escape decoding pass, and combining valid UTF-16 surrogate
pairs. Never recursively decode newly produced text. Reject unpaired surrogates
and NUL after decoding. Do not trim source, expand tabs or normalize Unicode.

For example, literal source `_x000D_` is serialized as `_x005F_x000D_` and decoded
back to seven literal characters. Actual form-feed is serialized as `_x000C_`
and decoded to one character. XML text `&<` is serialized as `&amp;&lt;`.

Canonical B content uses LF. Unexpected CR/CRLF inside B is not the file's EOL
setting; it is a recovery condition in section 9.

## 6. Exact reconstruction, splitting and hashes

### 6.1 Reconstruction algorithm

Given decoded blocks, their Gap values and a file's manifest metadata:

```text
GROUP complete physical part sequences into logical blocks
JOIN each block's B values without separators
ATTACH method blocks to the preceding class and check Parent
EXPAND class markers with method contents in order
START with empty file text
FOR EACH top-level block in order:
    BEFORE the first block, append Gap LF characters
    BEFORE every later block, append Gap plus one LF characters
    APPEND the expanded block content
APPEND exactly EndNL LF characters, even if there are no blocks
IF EOL is CRLF, replace every LF with CRLF
ENCODE as UTF-8; prepend EF BB BF only for utf-8-bom
```

Gap and EndNL are nonnegative integers. Missing Gap defaults to zero on the first
top-level block and one on later top-level blocks. Method Gap is ignored. Writers
SHOULD emit explicit Gap on first top-level parts to make these defaults visible.

Gap counts leading LF for the first block. Between later blocks it counts empty
lines **in addition to** the single separating LF. Thus `A` with Gap 0 followed
by `B` with Gap 1, EndNL 0, produces `A\n\nB`, not `A\nB`. Here and in all
vectors, `\n` denotes one actual LF. With Gap 0 for B, the result is `A\nB`.

### 6.2 Packing the inverse representation

Pack MUST decode supported source bytes strictly, detect BOM and EOL, normalize
CRLF to LF internally, then remove and count **all** trailing LF as EndNL.
Split the remaining body into blocks, preserving its characters exactly.

The baseline text splitter is language-independent:

- Empty boundary lines become Gap; an empty line contains zero characters.
  Whitespace-only lines remain source content.
- Use maximal runs of nonempty lines, divided into at most 200 lines per block.
  Consecutive subdivisions of a run have Gap 0.
- A block's content is its original lines joined with LF, without an added final LF.
- A suggested signature is `#TEXT-n` plus a trimmed first-line preview limited
  to 50 UTF-16 units; signatures do not affect file bytes.
- An empty body has no blocks. A zero-byte file therefore has no blocks and
  EndNL 0; a file consisting of two LF characters has no blocks and EndNL 2.

A syntax adapter MAY replace this layout with functions, types or other complete
source ranges. It MUST preserve all intervening text, indentation, comments and
line separators through Content and Gap. Nested or partial-line constructs that
cannot be isolated losslessly remain in surrounding text. Method extraction MUST
produce matching markers and method order. No particular AST library is required.

If optional analysis is unavailable or fails, Pack MUST retain an available text
path. Report fallback, or fail only when the caller explicitly required syntax-aware
processing. Never execute source to discover blocks. Every packed file MUST pass
an in-memory byte comparison against the bytes originally read before publication.

### 6.3 Chunking into physical parts

Measure every part in UTF-16 units. ChunkSize defaults to 30000 and must be in
1000..32000, below the cell limit.

1. If a logical block fits, emit one part, including for empty block content.
2. Otherwise walk LF-delimited lines, including each line's terminating LF.
   Greedily collect whole lines until the next would exceed ChunkSize; emit the
   collected part and continue. Do not preferentially split at empty lines.
3. For a line longer than ChunkSize, emit any pending collection, then split the
   complete overlong line into parts of at most ChunkSize. Keep surrogate pairs
   together; shorten a boundary by one unit if needed. Emit even a short final
   remainder of that line as its own part. Set `hard-split` for the file.
4. Preserve every LF within its part. Joining the parts must reproduce the exact
   logical block. Write Part, display signatures, part hashes and method Parents.
5. Set manifest Parts to the total number of physical block rows. Continuation
   rows have no Gap or first-part-only metadata.

Do not recalculate current row addresses by re-chunking when reading or selecting
blocks. Re-chunking occurs only on a write; it may change proposed row addresses.

### 6.4 Hash contracts

All hashes are SHA-256 written as 64 lowercase hexadecimal digits:

| Location / operation | Exact hash input |
|---|---|
| File-sheet F | UTF-8 without BOM of that physical part's decoded LF B value |
| Logical-block hash / ExpectedHash | UTF-8 without BOM of all part contents joined, in LF form |
| Manifest H / file hash | Exact reconstructed file bytes, including BOM, EOL conversion and EndNL |

Block hashes exclude Gap, EndNL, signatures, position, modes and file encoding.
A class block hash covers its **template**, not its separate methods. Expanded
class reads retain that template hash and template UTF-16 length even though the
returned Content includes methods. Changing a method changes method/file hashes
but not the template hash. A class hash precondition alone does not protect its
methods or their removal; callers must serialize access and check relevant methods
or a separately supported whole-file precondition.

Hashes detect differences, not intent, authorship or safety. Verify reports stale
hashes; explicit Fix accepts current readable contents and refreshes them. It
cannot recover lost source or establish that an external edit was intended.

## 7. Core operations

The following contracts are language-independent. Implementations must document
parameter spelling, input transport and result serialization separately, without
changing these meanings.

### 7.1 Pack

Inputs: source directory, output workbook, optional project name, ChunkSize,
exclusions, replacement permission and skip-unsupported policy.

- Read included regular files without following source links. Reject included
  symlinks/reparse points and special files, even when skipping unsupported text.
- Exclude the output workbook, its temporary artifacts and version-control
  metadata by default. Document matching rules and whether explicit exclusions
  replace or extend defaults.
- Preserve genuinely empty source directories. State whether filtering-created
  empty directories are also recorded; never promise recovery of excluded entries.
- Reject unsupported encoding/EOL/binary content by default. With an explicit
  skip policy, list skipped paths and reasons as expected outcomes, not warnings.
  I/O failures and unsafe links are not silently skippable text-format failures.
- Existing output requires explicit replacement permission. Input files should be
  quiescent or read from a snapshot; separate file reads are not a project-wide
  filesystem snapshot.
- Perform the per-file inverse/reconstruction comparison and the transactional
  workbook read-back in section 11.

Return output location, file/directory/sheet counts, skipped entries and issues.
A successful result describes the included subset, not all files in the source.

### 7.2 Unpack

Inputs: workbook, destination, optional Clean and DryRun. Reconstruct exact bytes
without requiring source-language runtimes. Preflight the full destination tree
before mutations as described in section 11.2.

Create recorded directories and file parents, write changed bytes, retain identical
file bytes, and restore recorded file modes where supported. Mode restoration is
also a mutation and MUST NOT occur in DryRun. Report unsupported mode restoration
or failures separately; do not claim exact mode preservation when it was impossible.

Clean MAY remove unrepresented files and unnecessary empty directories. It MUST
NOT follow links and MUST protect `.git` directories and descendants from deletion.
This deletion protection does not prohibit explicit project paths from writing
there; Unpack is not a security boundary for arbitrary executed project code.
A directory-wide export need not be atomic; report partial failure accurately.

Return actions per file (`created`, `updated`, `unchanged`, `deleted`) and relevant
directory/mode diagnostics. DryRun computes proposed actions without creating a
root, changing permissions, writing bytes or deleting anything.

### 7.3 Verify

Verify is read-only and MUST validate the project representation independently of
source language: package relationships, supported version, schemas, paths/tree,
required source fields, part sequences, Parent/marker consistency, character and
cell limits, reconstruction, declared counts and stored hashes.

It MUST NOT parse, compile, execute or install the stored source code, probe for
language runtimes or invoke language-specific subprocesses. A syntactically invalid
program may still be a structurally valid workbook. Source syntax diagnostics,
if desired, belong to a separate explicitly requested capability.

Return file count, issues, severity counts and a success/failure status. Diagnostics
SHOULD identify the project path, sheet and actual physical row when applicable.
Structural errors fail verification; warnings and informational hash changes do
not, unless a caller explicitly selects a stricter acceptance policy.

Verify SHOULD warn about logical blocks above 20000 UTF-16 units, cells at the
32767-unit limit and physical parts above configured ChunkSize. These are length
advisories, not proof of truncation or invalid source syntax. Hard cell limits
remain errors. Fix may re-chunk fully available B content, but cannot recover text
already lost by another tool; other oversized fields cannot be repaired by chunking.

Verify MAY compute block or file indexes but MUST NOT rewrite metadata or accept
new hashes. An ordinary workbook without the project layout does not pass merely
because its ZIP/XML is readable.

### 7.4 Fix

Fix is an explicit mutating operation, not an automatic consequence of Verify.
It accepts the current reconstructible contents as the new baseline and rebuilds
safe derived metadata. It MUST remain independent of source-language runtimes.

Fix MUST:

- Validate format version and reject unrecoverable errors before replacement.
- Reconcile H1 paths and manifest entries; report new and deleted paths.
- Keep encoding/EOL/Mode/EndNL for exact matching paths; use disclosed defaults
  for new paths only under the recovery rules.
- Preserve logical block boundaries and source contents. It MUST NOT reparse
  declarations, infer new block boundaries or split multiple definitions in a cell.
- Re-chunk existing logical blocks as necessary, rebuild headers, Parts, Parent,
  hashes and Flags, and retain valid existing sheet names.
- Preserve unknown configuration and extension data under section 10's policy.
- Write/read back transactionally, then run full **format** verification again.
  Warnings whose causes remain must remain in the final report.

It MUST NOT silently update hashes after a bootstrap failure in order to declare
that failed extraction successful. Review intended changes or a source diff before
accepting a new baseline.

## 8. Optional block-editor operations

### 8.1 Addressing and reads

Select a file by exact project-relative path, resolved through manifest/H1. Select
a block by its logical signature or the actual first-part row. If both are
provided, the explicit row wins. Normalize selector whitespace only, never content.
Use case-insensitive signature comparisons. A short form MAY omit a descriptive
type prefix or a method parameter list. Ambiguous matches MUST fail and list the
matching signatures with current physical rows; never silently choose the first.

List/Get results MUST provide path, signature, type, actual first row, part count,
UTF-16 content length, logical hash and owning class when present. List omits
content. Get joins physical parts without changing EOL from LF. A class read is
a template unless explicitly expanded; its hash/length retain template scope.
WholeFile returns exact reconstructed file bytes, including BOM and EOL. A textual
view MAY omit a BOM but MUST distinguish that from an exact-byte output.

### 8.2 Set, Add and Remove

Each mutation MUST support DryRun and the transaction/precondition rules below.
File-path creation, rename and deletion are separate operations or manifest/H1
reconciliation, not implicit block mutations.

| Operation | Contract |
|---|---|
| Set | Replace one block's Content; retain Gap and file metadata unless explicitly changed |
| Add | Insert a uniquely named block before/after a selected anchor, or at an explicitly selected end position |
| Remove | Remove one block; removing a class includes its methods; removing a method also removes its marker |

For textual Set/Add input, normalize transport CRLF/lone CR to LF, then remove
outer LF characters. Reject empty resulting content, NUL and unpaired surrogates;
preserve all spaces, tabs and internal LF. Input-file bytes, direct content and
stream input need a documented precedence and encoding contract. Do not rely on
host-specific pipeline newline behavior without declaring it.

Without a syntax adapter, the caller supplies type and signature for Add and any
explicit new navigation signature for Set. These metadata are not proof of source
grammar. Add/Set MUST reject new exact signature collisions. A syntax adapter MAY
validate one declaration, enforce source-name coherence or infer signatures, but
that remains optional and is not part of Verify/Fix.

Set on a class requires a template with the same number of method markers.
Class/method insertion must retain marker order and source indentation. Adding a
method requires an explicit template insertion location or a language adapter
that can find one safely; a generic editor MUST NOT assume a particular closing
brace, indentation rule or declaration syntax. If placement is ambiguous, fail
without mutation. Removing a marker MAY remove one adjacent empty line, but must
report that deliberate formatting policy.

An anchor that is a method, when used for ordinary top-level insertion, identifies
the entire owning class group. A suggested Add policy inserts one empty line
before later top-level blocks; a suggested Remove policy transfers the removed
block's Gap to its successor. Tools MUST document their policy instead of claiming
byte preservation for the deliberately added/deleted region.

Renaming a class must consistently handle child navigation signatures if they
encode its name; otherwise reject the rename. Do not leave silently stale child
names. Existing block positions are not stable after part growth, insertion or
removal; return the new row and distinguish it from the selected original row.

### 8.3 Hash preconditions and results

Set/Remove ExpectedHash compares the selected logical block hash. Add ExpectedHash
compares the explicitly selected Before/After anchor. Trim surrounding input
whitespace and compare lowercase hexadecimal values. A mismatch aborts before any
write. Reject an unanchored Add hash precondition unless a separately specified
file/class/append precondition is actually implemented and checked.

A Force option MAY bypass optional syntax checks; it MUST NOT bypass character,
marker, uniqueness, hash, path or read-back checks. Confirmation is an application
choice. If provided, there must be one documented end-to-end confirmation boundary,
not duplicate prompts or silently ignored requests.

Mutation results SHOULD identify operation, path, old/new logical hashes, file
hash, current/new row, new part count and total file rows. Remove has no replacement
block/hash. DryRun reports proposed values and Changed false; persisted success
reports Changed true. Human-readable summaries are not machine-stable identifiers.

## 9. Verification severities and controlled recovery

A conforming reader MUST read complete generated workbooks and apply defined
field defaults. Core Fix MUST support the unambiguous manifest reconciliation and
derived-field refresh described in section 7.4. Other recovery is optional and
must be declared explicitly. Recovery MUST report assumptions and cannot claim
knowledge of original BOM/EOL/newline counts that were never stored. Unsupported
recovery must fail clearly rather than guess.

| Condition | Normal verification / permitted explicit repair |
|---|---|
| Unsupported FormatVersion | Error; no Fix or downgrade |
| Missing config or manifest | Warning during recovery inspection; explicit Fix may generate using disclosed defaults and valid H1 paths |
| Missing instructions | Advisory; generate on full export/Fix |
| Invalid H1/path/tree, duplicate authoritative mappings or config keys | Error; no guessed identity or silent overwrite |
| Formula in source/authoritative cells, NUL or unpaired surrogate | Error; no automatic content repair |
| Unknown block type or broken/incomplete part sequence | Error; no missing-content recovery |
| Missing/inconsistent fixed headers | Warning if fields remain unambiguous; Fix rewrites headers |
| Stale Parent | Error in strict reconstruction; Fix may regenerate only from unambiguous adjacency and valid method/marker order, never by moving methods |
| Declared Parts differs from actual physical count | Warning; Fix recomputes |
| Missing, malformed or stale derived hash | Warning for missing/malformed values, Info for changed values; Fix recomputes without asserting intent |
| CR/CRLF inside B | Noncanonical content; explicit recovery may normalize to LF with a warning, never infer file EOL from it |
| Invalid Gap/EndNL/encoding/EOL/mode | Error in strict reading; recovery may use the documented new-file defaults with a warning and loss-of-original-metadata notice |
| Missing Gap alone | Apply the defined defaults: 0 for the first top-level block and 1 for later blocks; Fix may persist them explicitly |
| More methods than markers | Recovery may add markers only with an explicitly supplied or unambiguously determined template insertion point; no assumed language delimiter |
| More markers than methods | Error; missing method content cannot be invented |
| File sheet without matching manifest path | Info during reconciliation; Fix adds defaults and reports metadata loss |
| Manifest file without a corresponding H1 path | Warning; Fix removes the stale entry, without claiming source recovery |
| Duplicate block label | Info; reconstruction can remain valid, but selection must detect ambiguity |
| Oversized cell/block or portability issue | Error for hard format limits; Warning for legal but suspicious lengths or host portability |
| Source syntax error | Outside Verify/Fix; only an optional syntax capability may report it |

When replacing missing/invalid ChunkSize in explicit recovery, use 30000 and report
it. Retain a valid ProjectName; otherwise use the workbook basename. Unsupported
FormatVersion must never use this defaulting rule.

Part hashes and file hashes are required in complete writer output but may be
stale/absent in a workbook undergoing external edits. This distinction explains
why Verify can report a changed hash without failing while a bootstrap reader
must reject it as an unverified input. The format and hash input do not change.

## 10. Preservation contracts

### 10.1 Full export and annotations

Pack regenerates the workbook from source files. Fix regenerates derived workbook
structures while preserving reconstructed file bytes and existing logical layout.
Neither operation promises identical ZIP bytes, Excel formatting or unknown Office
features. Unknown auxiliary sheets MUST survive at least as text values with
addresses in a declared full-export profile; otherwise reject unsupported content
or explicitly obtain approval for the documented loss. Do not silently advertise
arbitrary spreadsheet-feature preservation.

Extra file-sheet columns J onward are annotations/extensions outside reconstruction.
A preserving editor MUST retain their headers and values, including continuation
rows. The default ownership model attaches each part's annotations to that ordinal
part of the logical block. Re-chunking may preserve them at the same ordinals;
if shrinking would remove nonempty annotations, fail unless an explicit relocation
or discard policy has been selected. Intentional deletion of a block may delete
its associated annotations as part of that operation.

New mandatory reconstruction fields or unknown type semantics require a separately
specified extension/version; annotation columns cannot secretly alter file bytes.

### 10.2 Targeted package updates

A block editor SHOULD parse only the common metadata and affected file sheet,
then copy unrelated uncompressed ZIP-part payloads unchanged. It may regenerate
target `sheetData`, required row-bound/range metadata and the matching manifest
row. Preserve unrelated outer sheet XML where possible; rewritten cells may lose
rich-text formatting without changing decoded values.

If shared strings exist, retain existing entries and indexes, appending new texts.
Unused entries are allowed. Without a string table, inline-string editing is valid.
Do not renumber shared strings without updating every reference. If reporting
orphans, count distinct referenced indexes rather than total references.

Update the affected file's hash, Parts and Flags. If its manifest entry is missing,
a targeted editor MAY leave it absent with an explicit recovery-required result;
otherwise it must reconcile through the Core Fix rules, not guess metadata.
Read-back MUST confirm affected file bytes and promised untouched ZIP payloads.

A targeted editor assumes untouched file sheets were previously verified. After
external spreadsheet changes, run whole-project Verify before relying on that
fast path. Full Fix may compact strings, but MUST NOT reparse source or change
logical block boundaries merely to resemble a newly packed workbook.

## 11. Write safety and filesystem behavior

### 11.1 Workbook transactions

Pack, Fix and block mutations MUST stage a temporary workbook beside the output,
validate it and replace the destination only on success:

```text
READ the input model and capture the original workbook identity when applicable
VALIDATE version, structure, preconditions and intended reconstruction
BUILD the proposed workbook in memory
WRITE a temporary workbook in the destination directory
READ it back through the same format reader
COMPARE path sets, file counts, reconstructed bytes/hashes and empty directories
CHECK layout/annotations and untouched payloads to the promised preservation scope
RECHECK that the original workbook has not changed since it was read
REPLACE the destination atomically using the host's supported facility
CLEAN UP temporary artifacts on all pre-replacement failures
```

For replacement, compare an original-content digest or equivalent exact content
snapshot, not only a block hash. Refuse unexpectedly appearing output when creation
was requested. Recheck/replace can still have a race unless the host supplies a
stronger exclusive mechanism; document that limit. One writer at a time is the
baseline coordination rule. Organizational locking is not a format feature.

Failures before replacement MUST leave the previous workbook intact. Preserve
workbook container permissions/ownership/ACLs according to the host application's
policy; they are separate from stored project file modes. Atomic rename alone
is not a claim of power-loss durability. State any additional durability guarantees.

### 11.2 Destination containment and previews

Before Unpack writes, preflight all requested file/directory paths and relevant
existing destination entries. Reject unsafe symbolic links/reparse points at the
root, existing ancestors or final targets; lexical containment alone is insufficient.
Clean traversal MUST NOT follow links outside the destination.

Use host facilities that prevent redirected traversal, or require exclusive
control of a freshly created destination. If a platform cannot meet this contract,
restrict that mode or report it unsupported; do not silently weaken containment.
Another process moving already opened directories is outside an exclusively
controlled-tree assumption and must be considered when stronger isolation is needed.

Unpack and block mutations MUST provide a no-write DryRun. Other mutating operations
SHOULD provide one. A preview may validate and calculate proposed results but MUST
NOT create directories or temporary files, change modes, write source/workbook
bytes or delete entries. A preview is not a persisted read-back transaction.

### 11.3 Versions, external writers and transport

Keep a prior version before accepting external spreadsheet edits or Fix. Read back
complete intended values, verify and inspect reconstructed source diffs. Hashes do
not distinguish intentional edits from truncation. Synchronization must be quiescent
during local replacement; a neighboring application lock file may warrant a warning
but is not a distributed lock or a guarantee that unsaved edits do not exist.

Version-control review and merge SHOULD use unpacked source files; after resolving
source conflicts, pack and verify again. Remote transports, version tokens and
conflict-copy detection belong to separate integrations. They must not be implied
by successful verification of one local workbook.

## 12. Optional syntax analysis and execution

A syntax adapter MAY provide language-specific block extraction, edit validation,
name inference and safe class insertion positions. It must report the language,
capability and dependencies being used. Core Verify/Fix MUST NOT invoke it.

For example, an implementation may analyze Java, Python or PowerShell declarations
with an appropriate parser while another implementation stores those same files
as text blocks. Both can conform to the same format and reconstruct identical bytes.
Optional parser checks MUST NOT require evaluating the source. Parser-relative
locations should be mapped back to supplied block/source lines when possible.

An Executor MAY reconstruct into a fresh temporary project, select an explicitly
requested existing project-relative entry point, set the project root as working
directory and invoke a caller-selected tool/runtime. Preserve argument order,
forward output and return the child status. Clean up in all completion/failure
paths unless retention was requested. Compilation or dependency installation needs
its own explicit contract; it is not inferred from a Language value.

Execution is ordinary project-code execution, not a sandbox. It must never happen
as a side effect of reading, Pack, Unpack, Verify, Fix or hash checking.

## 13. Independent bootstrap

An agent given only this document and a complete generated workbook can implement
a bootstrap reader with ZIP, XML, Unicode/UTF-8, hashing and filesystem primitives.
No installed copy of the packaged tools is required. A convenience spreadsheet
library is optional; its string decoding must pass the vectors below.

Bootstrap is intentionally narrower than repair:

- Require supported format, authoritative metadata, complete part sequences and
  unambiguous paths. Use a fresh destination under an agent-controlled parent.
- Reconstruct using sections 2 through 6, without source syntax analysis or execution.
- Per-part hash checking is optional for this profile. Comparing each saved file's
  exact bytes against manifest H is mandatory before reporting success or executing
  recovered code. Missing/mismatching file hashes stop handover; do not Fix them.
- Report extracted and verified file counts from the actual manifest, plus any
  skipped optional tests or unavailable host mode support. Do not hard-code counts.
- After inspecting recovered code, use its documented package/application entry
  point and import/load paths. Launching a package's internal file directly may
  not be equivalent to starting the package. Isolated runtime modes may ignore
  environment-based load paths; configure those paths explicitly if needed.

```text
OPEN the workbook as data and resolve its package relationships
DECODE original strings without lossy or repeated escape processing
READ configuration and manifest; validate version and destination paths
FOR EACH included file:
    ASSEMBLE physical parts and expand method templates
    RECONSTRUCT exact bytes using Gap, EndNL, EOL and Encoding
    WRITE inside the freshly controlled destination
    COMPARE saved file SHA-256 with manifest H
CREATE recorded empty directories and restore supported file modes
REPORT actual counts and any failures or limitations
HAND OVER inspected sources only after every file hash agrees
```

## 14. Compatibility vectors and acceptance criteria

### 14.1 Exact bytes and hash inputs

In these data vectors `\n`, `\r` and `\t` denote actual LF, CR and tab. Unless
specified otherwise: UTF-8 without BOM, LF, no methods. `(Content, Gap)` describes
a top-level logical block. Content is already decoded Unicode.

| ID | Blocks / metadata | Expected file bytes, hex |
|---|---|---|
| V1 | No blocks; EndNL 0 | Empty byte sequence |
| V2 | No blocks; EndNL 2 | `0a0a` |
| V3 | `(A,0)`, `(B,0)`; EndNL 0 | `410a42` |
| V4 | `(A,0)`, `(B,1)`; EndNL 1 | `410a0a420a` |
| V5 | V4 with UTF-8 BOM and CRLF | `efbbbf410d0a0d0a420d0a` |
| V6 | `(A,2)`; EndNL 0 | `0a0a41` |
| V7 | One block from parts `A\n` and `B`; Gap 0; EndNL 0 | `410a42` |
| V8 | `(A,0)`, EndNL 0; add a new empty file sheet using defaults | Original file `41`; new file `0a` |
| V9 | No blocks; EndNL 0; UTF-8 BOM | `efbbbf` |
| V10 | `(A,0)`, `(B,1)`; EndNL 0 | `410a0a42` |
| V11 | `(A,1)`; EndNL 0 | `0a41` |
| V12 | One block containing space then tab; Gap 0; EndNL 0 | `2009` |

Known file SHA-256 values are:

| Vector | SHA-256 |
|---|---|
| V1 | `e3b0c44298fc1c149afbf4c8996fb92427ae41e4649b934ca495991b7852b855` |
| V2 | `75a11da44c802486bc6f65640aa48a730f0f684c5c07a42ba3cd1735eb3fb070` |
| V3 | `23519a43c66b4c342f25b32e09797ec5f3fc0be388cd8243fb3449afbdce4013` |
| V4 | `c410bfd99c846b9dc84431c34ca7044a1e3e4e33146a8fa105b25577e6188768` |
| V5 | `93585c2d3e1d7b051d33647d5b1d922a21d5d599c819be3958c59fadb49f881e` |

Class vector: template `TYPE K\n#@METHOD@\nEND` and method content
`    METHOD M -> 1`, Gap 0, EndNL 1, reconstruct
`TYPE K\n    METHOD M -> 1\nEND\n`. This is illustrative text, not a required
source grammar. Replacing the method's `1` with `2` changes method/file hashes
but leaves the template hash unchanged. Expanded Get retains template hash/length.

### 14.2 XML, escapes and Unicode

These inputs are original XML text spellings, before XML entity decoding:

| Input | Expected decoded value | UTF-8 bytes, hex |
|---|---|---|
| `_x005F_x000D_` | Literal `_x000D_`, no CR | `5f78303030445f` |
| `_x005F_x005F_` | Literal `_x005F_` | `5f78303035465f` |
| `_x005F_x005F_x005F_x000C_` | Literal `_x005F_x000C_` | `5f78303035465f78303030435f` |
| `_x000C_` | One form-feed | `0c` |
| `_x000D_` | One CR | `0d` |
| `_xD83D__xDE00_` | One emoji, 😀 | `f09f9880` |
| `&amp;&lt;` | `&<` | `263c` |
| `_xD800_` | Error: unpaired surrogate | No output |
| `_x0000_` | Error: forbidden NUL | No output |

A decoded CR is a valid adapter result but noncanonical within B; do not confuse
it with a literal `_x000D_` or the file's CRLF metadata. Separate rich-text runs
`_x005F_` and `x000D_` join before decoding to produce literal `_x000D_`.
Phonetic runs do not become source text; empty shared/inline strings remain empty.
All valid vectors MUST survive writing and reading without recursive decoding.

ChunkSize 1000, Content of 999 ASCII `a` followed by one emoji: part lengths are
999 and 2 UTF-16 units, `hard-split` is set, and joining produces 1001 units.
For 1000 `a` followed by LF, part lengths are 1000 and 1; the LF belongs to the
long line. No unpaired surrogate may be emitted at any boundary.

### 14.3 Required implementation checks

An implementation MUST publish its supported profiles and qualify their behavior.
The following are acceptance criteria, not predeclared evidence of any tool's status:

| Area | Required checks |
|---|---|
| Byte round trip | V1–V12; BOM-only, LF/CRLF, absent/multiple final newlines, leading empty lines, whitespace-only lines, tabs, Unicode and permitted controls |
| Splitting/chunking | Paragraphs beyond 200 lines, exact UTF-16/cell boundaries, overlong lines, complete ordered parts and correct Parent/Gap/Flags |
| Optional syntax adapter | Complete declarations, nested/conditional/partial-line constructs, comments, constructors and retained text; fallback or explicit failure preserves bytes |
| Classes/methods | Templates versus Expanded reads, marker order/count, method addition/removal, class removal and child-signature rename policy |
| Core verification | Mixed source languages and invalid source syntax with no runtime lookup, parser invocation, compilation or subprocesses; Verify and Fix must still validate format |
| Hash preconditions | Wrong block/anchor hash leaves workbook unchanged; unsupported preconditions are rejected, not ignored |
| Fix/recovery | Each supported section 9 case has declared severity/repair behavior; logical boundaries and surviving advisories remain; no version downgrade |
| Invalid input | Duplicate authoritative identities, infeasible trees, unsafe paths, broken parts, formulas, NUL and unpaired surrogates are diagnosed before publication |
| Destination safety | Root/ancestor/final symlinks and Clean cannot alter an outside sentinel; host collisions fail before writing |
| DryRun | No absent-root creation, bytes, modes, temporary artifacts or Clean deletions; proposed results are clearly unpersisted |
| Transaction failure | Inject validation, reconstruction/read-back and replacement failures; previous workbook remains intact and temporary files are cleaned before replacement |
| Concurrency | Changed original or newly appearing output aborts; documented race limits and coordination assumptions are explicit |
| Preservation | First-part and continuation annotations/headers, auxiliary text values, shared/inline/rich strings and promised untouched payloads survive; unsupported feature losses are disclosed |
| Addressing/results | Actual physical row addresses survive sparse/noncanonical input; ambiguity reports matches; post-edit rows are distinguished from original rows |
| Bootstrap | Fresh reader independent of packaged tools reconstructs and verifies all files; no source runtime or execution is needed |
| Platforms/libraries | Test each claimed host and external writer with dependencies installed; report skipped checks separately from passes |

Conformance is demonstrated by reconstructed bytes, logical values and the declared
preservation scope. ZIP-byte equality is required only for a separately claimed
container-identity guarantee. No programming language or implementation is the
normative authority for these results.
