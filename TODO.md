# TODO — `xls2dat`

Remaining backlog after the .NET 8 rewrite. Items are tagged **[BUG]**, **[ROBUST]**, **[PORTABLE]**, **[ENHANCE]**, or **[CHORE]**.

## Done in this rewrite

- ✅ Cross-platform: replaces Windows-only Excel COM Interop with [ClosedXML](https://github.com/ClosedXML/ClosedXML) (native `.xlsx`/`.xlsm`) and LibreOffice headless (`.xls` / `.ods` / `.numbers`).
- ✅ Concordance encoding bugs from the old script — TAB→delimiter conversion ([src/Xls2Dat.Core/Formatting/DatLineFormatter.cs](src/Xls2Dat.Core/Formatting/DatLineFormatter.cs)), UTF-16 LE w/ BOM output ([src/Xls2Dat.Core/Writers/ConcordanceDatWriter.cs](src/Xls2Dat.Core/Writers/ConcordanceDatWriter.cs)), embedded newline → `®` (no more silent data loss).
- ✅ Removed the `Stop-Process -Name "EXCEL"` footgun — no more killing the user's other Excel windows.
- ✅ Configurable delimiters via `--field-delimiter`, `--text-qualifier`, `--newline-replacement`.
- ✅ Format detection by extension *and* magic bytes ([src/Xls2Dat.Core/Detection/FormatDetector.cs](src/Xls2Dat.Core/Detection/FormatDetector.cs)).
- ✅ Test suite: 29 xunit tests covering formatter, detector, CSV reader, writer.
- ✅ Repo hygiene: `.gitignore`, `.gitattributes`.

---

## 1. High-priority follow-ups

### 1.1 [ENHANCE] OpenXmlReader streaming
- ClosedXML loads the workbook fully into memory. For multi-GB `.xlsx`, swap the ClosedXML reader for a SAX-style `DocumentFormat.OpenXml` implementation that yields rows without materializing the full sheet.
- File: [src/Xls2Dat.Core/Readers/OpenXmlReader.cs](src/Xls2Dat.Core/Readers/OpenXmlReader.cs).

### 1.2 [ENHANCE] Batch / glob input
- Today the CLI takes one input file. Accept multiple inputs or a glob:
  `xls2dat ./data/*.xlsx -o ./out` — convert every file in parallel, writing to the same out dir.

### 1.3 [ROBUST] Integration test against a real `.xlsx` fixture
- The current OpenXmlReader has no test. Generate a small `.xlsx` (programmatically via ClosedXML in a test fixture) covering multi-sheet, multi-type cells (date, number, formula, empty), embedded newline, and Unicode — then assert the produced `.dat` bytes.

### 1.4 [ROBUST] LibreOffice integration test (opt-in)
- Add a `[Trait("category", "libreoffice")]` test that converts a `.ods` fixture end-to-end. Skip when `soffice` isn't on PATH so CI can opt-in.

### 1.5 [ENHANCE] CI pipeline
- GitHub Actions workflow: build + test on `windows-latest`, `ubuntu-latest`, `macos-latest`. Publish artifacts on tag push.

### 1.6 [BUG] CSV `BadDataFound = null` swallows malformed input silently
- [src/Xls2Dat.Core/Readers/CsvReader.cs](src/Xls2Dat.Core/Readers/CsvReader.cs) suppresses CsvHelper's bad-data callback. Surface it as a warning to stderr in `--verbose` mode at minimum.

---

## 2. Format coverage

### 2.1 [ENHANCE] Native `.ods` reader
- LibreOffice fallback is slow (spawns a subprocess per file) and adds an install dependency. A native ODS reader (e.g. via [NPOI](https://github.com/nissl-lab/npoi) or a custom `.ods` zip+xml parser) would eliminate the dependency for the common case.

### 2.2 [ENHANCE] Native `.xls` (legacy BIFF) reader
- NPOI's `HSSFWorkbook` reads `.xls` natively, cross-platform, no LibreOffice required.

### 2.3 [ENHANCE] Native `.numbers` reader
- Realistically very hard — the format is undocumented Protobuf. Keep the LibreOffice route unless an upstream library appears.

---

## 3. Output / DAT semantics

### 3.1 [ENHANCE] Column inventory / metadata sidecar
- Concordance load-files often pair `.dat` with an `.opt` for opticons and a `.log` for warnings. Emit a `.log` summarizing per-sheet counts, skipped rows, and any field-truncation events.

### 3.2 [ENHANCE] Field-width / length warnings
- Concordance has historical max-field-width limits (~32 KB). Emit a warning when any field exceeds a configurable threshold (`--max-field-bytes`).

### 3.3 [ENHANCE] Header validation
- Optional `--header-schema <file>` switch that compares the first row to a YAML/JSON-defined column list and fails fast on mismatch.

### 3.4 [ENHANCE] Date / number formatting policies
- ClosedXML's `GetFormattedString()` honors the cell's display format. Expose `--date-format` / `--number-format` overrides for deterministic output across locales.

---

## 4. Repo & ops chores

- **[CHORE]** Add a `LICENSE` reference link from README (the file exists at repo root).
- **[CHORE]** Add `CHANGELOG.md` and start tagging releases.
- **[CHORE]** `dotnet format` config / `.editorconfig`.
- **[CHORE]** Document the LibreOffice install on each platform in README (Homebrew, apt, winget commands).
- **[CHORE]** Publish to NuGet as a `dotnet tool`:
  `dotnet tool install -g xls2dat`.
- **[CHORE]** Add `--version` output (read from assembly informational version).
