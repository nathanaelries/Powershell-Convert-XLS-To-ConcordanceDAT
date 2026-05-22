# TODO — Bug Fixes, Robustness & Platform Portability

A prioritized backlog for the PowerShell XLS→Concordance DAT converter (`src/`) and the bundled C# `VenioNSRLTool`. Items are tagged **[BUG]**, **[ROBUST]**, **[PORTABLE]**, **[ENHANCE]**, or **[CHORE]**, and ranked by likely impact.

---

## 1. Critical correctness bugs (PowerShell converter)

### 1.1 [BUG] Field delimiter conversion is missing
- `Format-DATLine` ([src/utils/file-converter.psm1:132](src/utils/file-converter.psm1#L132)) operates on the **PIPE** character (`[char]0x14`), but Excel's `xlUnicodeText` SaveAs writes **TAB-delimited** UTF-16 files. Tabs are never converted to `0x14`, so the resulting `.dat` lacks proper column delimiters.
- `$CONFIG.TAB` is defined ([src/config/settings.psm1:6](src/config/settings.psm1#L6)) but never referenced.
- **Fix:** in `Format-DATLine`, replace `TAB` with `THORN + PIPE + THORN` before the existing wrapping logic, and stop wrapping `PIPE` (there is none on input).

### 1.2 [BUG] `Quit-Excel` kills *every* Excel process on the host
- [src/utils/excel-handler.psm1:19](src/utils/excel-handler.psm1#L19) calls `Get-Process -Name "EXCEL" | Stop-Process -Force`. If the user has any other workbook open, it is force-closed and unsaved data is lost.
- **Fix:** capture the PID from the COM object (`[System.Runtime.InteropServices.Marshal]::GetIUnknownForObject` + `GetWindowThreadProcessId`) and only stop that one. Prefer not to `Stop-Process` at all — `Quit()` + `ReleaseComObject` is sufficient when references are cleaned up.

### 1.3 [BUG] Excel `Quit()` called twice on the happy path
- [src/main.ps1:85-86](src/main.ps1#L85-L86) calls `$excel.Quit()` then `Quit-Excel $excel` (which calls `Quit()` again). Throws a COM exception in the `finally` block that masks earlier errors.
- **Fix:** remove the redundant `$excel.Quit()`; let `Quit-Excel` own the lifecycle.

### 1.4 [BUG] Encoding mismatch on the temp-read path
- Excel's `xlUnicodeText` writes UTF-16 LE with BOM, but [src/utils/file-converter.psm1:20](src/utils/file-converter.psm1#L20) reads `Convert-ExcelToTemp` with `Encoding.Default` (system ANSI). Multibyte / non-ASCII content is mojibaked.
- **Fix:** pass `Encoding.Unicode` (or let the BOM-detection in `StreamReader` work by passing `Encoding.Default` together with `detectEncodingFromByteOrderMarks = $true` and *trusting* the BOM — the latter is already enabled, but constructing the reader still defaults to ANSI when no BOM is present).

### 1.5 [BUG] Embedded newlines inside cells are destroyed
- [src/utils/file-converter.psm1:143](src/utils/file-converter.psm1#L143) does `$line -replace "\`n|\`r"`, which silently drops embedded CR/LF inside quoted cells, collapsing multi-line content onto a single line and corrupting downstream metadata.
- **Fix:** Concordance convention is to replace embedded LF inside fields with `®` (`0xAE`). Detect LF/CR *inside* quoted fields and substitute; only strip line terminators *between* records.

### 1.6 [BUG] `using module '..\config\settings.ps1'` is invalid
- [src/utils/excel-handler.ps1:2](src/utils/excel-handler.ps1#L2) and [src/utils/file-converter.ps1:2](src/utils/file-converter.ps1#L2) reference `.ps1` files via `using module`. `using module` requires a `.psm1`, `.psd1`, or module-folder path. These `.ps1` variants of the modules are dead code.
- **Fix:** delete the `.ps1` duplicates (see also §2.1) or convert callers to `.psm1`.

### 1.7 [BUG] `MaxThreads` parameter has no effect
- `src/main.ps1` advertises a `-MaxThreads` parameter but only sets `$CONFIG.MAX_THREADS`. The active conversion path (`Convert-TempToDAT` in `.psm1`) is single-threaded; `task-processor.psm1` is orphaned.
- **Fix:** either wire `task-processor.psm1` into `Convert-TempToDAT`, or drop the parameter and the unused modules until parallelism is actually implemented.

### 1.8 [BUG] `foreach` swallows worksheet failures
- [src/main.ps1:120-122](src/main.ps1#L120-L122) calls `continue` on error and the script still exits 0.
- **Fix:** track failure count, exit non-zero when any worksheet fails, and surface the per-sheet error summary at the end.

### 1.9 [BUG] Brittle `Start-Sleep -Milliseconds 100` after `SaveAs`
- [src/main.ps1:73](src/main.ps1#L73) — a 100 ms sleep after `SaveAs` looks like it papers over a race. Either remove it or document the COM behavior it's mitigating.

---

## 2. Robustness & code health (PowerShell converter)

### 2.1 [CHORE] Remove duplicated `.ps1`/`.psm1` modules
- `src/utils/excel-handler.ps1` vs `.psm1`, `src/utils/file-converter.ps1` vs `.psm1`, `src/config/settings.ps1` vs `.psm1`. The `.ps1` copies are stale and reference the buggy single-pass formatter. Keep one canonical set.

### 2.2 [ROBUST] Add Pester tests
- No tests exist. At minimum cover `Format-DATLine` (quoting, embedded delimiters, embedded newlines, empty fields, Unicode), `Get-OptimalThreadCount`, and `Get-ChunkSize`.
- Add fixture XLS/XLSX files and golden DAT outputs under `tests/fixtures/`.

### 2.3 [ROBUST] Add PSScriptAnalyzer config and run in CI
- No linter config. Several easily-caught issues exist (unused variables, mixed quoting, plural noun in `Remove-TempFiles` — should be `Remove-TempFile`, etc.).

### 2.4 [ROBUST] Replace `Write-Host` with `Write-Verbose` / `Write-Information`
- `Write-Host` bypasses output streams and prevents piping. Use `-Verbose` / `-InformationAction` for diagnostics; reserve `Write-Host` for explicit interactive output (or use `Write-Information -InformationAction Continue`).

### 2.5 [ROBUST] Validate input file format up front
- `main.ps1` accepts any path that exists. Add explicit handling for `.xls`, `.xlsx`, `.xlsm`, `.csv`. Reject others with a clear error before spinning up Excel.

### 2.6 [ROBUST] Detect password-protected workbooks
- Currently they hang Excel. Pass a sentinel password and catch the prompt, or check the workbook's protection state before opening.

### 2.7 [ROBUST] Output filename hygiene
- `$workbook.Name` includes the extension, so temp/output names become `Workbook.xlsx_Sheet1.dat`. Strip the extension before composing `tempName`.

### 2.8 [ROBUST] Configurable delimiters
- Expose `-FieldDelimiter`, `-TextQualifier`, and `-NewlineReplacement` parameters (defaulting to `0x14`, `0xFE`, `0xAE`) so the script can target other Concordance variants and EDRM-style outputs.

### 2.9 [ROBUST] Resumable / idempotent runs
- If the script crashes mid-sheet, temp files in `$env:TEMP` are orphaned. Use a deterministic temp subdirectory (e.g. `Join-Path ([IO.Path]::GetTempPath()) "xls2dat-$([Guid]::NewGuid())"`) and clean it on success/failure via `try/finally`.

### 2.10 [ROBUST] Streaming progress for large sheets
- `Convert-TempToDAT` calls `Write-Progress` every `BATCH_SIZE` rows but `-PercentComplete -1` produces an indeterminate bar. Compute real percent from total-line-count (read once) or use byte offsets.

### 2.11 [ROBUST] Header row option
- Add `-IncludeHeader` / `-SkipHeader` switches and validate the first row matches a supplied schema if provided.

---

## 3. Platform agnosticism (PowerShell converter)

### 3.1 [PORTABLE] Replace Excel COM Interop
- COM Interop requires Windows **and** an installed copy of Microsoft Excel. Replace with one of:
  - **`ImportExcel`** PowerShell module (EPPlus-based, cross-platform, MIT-ish licensing).
  - **ClosedXML** via `Add-Type` (cross-platform, .NET).
  - **OpenXML SDK** for direct streaming reads of `.xlsx` (best for very large files).
- Keep a `-UseExcelInterop` legacy switch for `.xls` (binary BIFF) files that the OpenXML libraries cannot read.

### 3.2 [PORTABLE] Use `[IO.Path]::GetTempPath()` instead of `$env:TEMP`
- `$env:TEMP` is not consistently populated on Linux/macOS PowerShell 7+ sessions.

### 3.3 [PORTABLE] Drop `Stop-Process -Name "EXCEL"`
- See §1.2 — also un-runnable on Linux/macOS.

### 3.4 [PORTABLE] Target PowerShell 7+
- Validate the script against `pwsh` 7.4 LTS in addition to Windows PowerShell 5.1. Add `#Requires -Version 7.2` (or 5.1 with conditional logic for COM).

### 3.5 [PORTABLE] CSV fast-path
- For CSV inputs, bypass Excel entirely and stream with `[IO.File]::ReadLines` + `Format-DATLine`. Works identically on every platform and is dramatically faster.

---

## 4. C# `VenioNSRLTool`

### 4.1 [BUG] SQL injection in `CREATE DATABASE` / `db_id`
- [VenioNSRLTool/Helpers/DatabaseHelper.cs:20](VenioNSRLTool/Helpers/DatabaseHelper.cs#L20) interpolates `dbName` into both `db_id('{dbName}')` and `CREATE DATABASE [{dbName}]`. A name containing `]`, `'`, or `;` breaks out.
- **Fix:** validate `dbName` against `^[A-Za-z_][A-Za-z0-9_]{0,127}$`, quote with `QUOTENAME()` server-side where possible, and parameterize `db_id` (`SELECT db_id(@n)`).

### 4.2 [BUG] Connection-string composition is unsafe
- `BuildConnectionString` ([VenioNSRLTool/Helpers/DatabaseHelper.cs:13](VenioNSRLTool/Helpers/DatabaseHelper.cs#L13)) string-concatenates user input. A password containing `;` or `"` breaks the string and can inject options.
- **Fix:** use `Microsoft.Data.SqlClient.SqlConnectionStringBuilder` with property setters; never concatenate.

### 4.3 [ROBUST] UI updates from async continuations
- `txtLog.AppendText(...)` is invoked after `await` in `RunNISTImport`. Depending on `SynchronizationContext`, this may execute off the UI thread.
- **Fix:** wrap log writes in `Invoke(() => txtLog.AppendText(...))` or use `IProgress<string>`.

### 4.4 [ROBUST] Hard-coded INI path
- `iniPath = @"C:\Program Files\Venio\VenioFPR\VenioSetup.ini"` ([VenioNSRLTool/MainForm.cs:13](VenioNSRLTool/MainForm.cs#L13)) blocks dev/test machines and non-default installs.
- **Fix:** allow overriding via CLI arg, env var (`VENIO_SETUP_INI`), or a config file; fall back to the default.

### 4.5 [ROBUST] Plain-text password held in field
- `sqlPassword` is a `string` field on `MainForm`. Use `SecureString` (or zero the array after building the connection string) and avoid keeping it in process memory longer than needed.

### 4.6 [ROBUST] Cancel/close handling on the password prompt
- `PromptForPassword` returns `tb.Text` regardless of how the dialog closes. Distinguish `DialogResult.OK` from `Cancel` and propagate cancellation.

### 4.7 [ROBUST] No retry/backoff on NIST downloads
- Network calls in `NetworkHelper.GetLatestNSRLVersion` should retry with exponential backoff and honor a configurable timeout.

### 4.8 [ROBUST] Verify SQL principal has required permissions
- Probe `HAS_PERMS_BY_NAME(null, 'SERVER', 'CREATE ANY DATABASE')` before showing the **Next** button; fail fast with a clear message rather than midway through an import.

### 4.9 [PORTABLE] WinForms restricts the tool to Windows
- Long-term: extract the database/NSRL logic into a `netstandard2.0` (or `net8.0`) library and provide:
  - a WinForms front-end for current users, and
  - a `dotnet`-based CLI (`venio-nsrl import --server ... --db ...`) for Linux/macOS and CI use.

### 4.10 [CHORE] Add the `.csproj` and a `dotnet test` project
- No project file is in the tree (or it isn't tracked). Add `VenioNSRLTool.csproj`, an `xUnit` test project, and a `.gitignore` covering `bin/` and `obj/`.

---

## 5. Repo-wide chores

- **[CHORE]** `.gitignore` for `bin/`, `obj/`, `*.user`, `.vs/`, `*.tmp`, `*.dat`.
- **[CHORE]** `.gitattributes` to pin `* text=auto eol=lf` (or `crlf` for `*.ps1` if preferred) — current repo has no line-ending policy.
- **[CHORE]** `.editorconfig` aligning indentation, trailing-whitespace, and final-newline rules across `.cs`, `.ps1`, `.psm1`, and `.md`.
- **[CHORE]** GitHub Actions workflow: PSScriptAnalyzer + Pester for `src/`, `dotnet build` + `dotnet test` for `VenioNSRLTool/`.
- **[CHORE]** Expand `README.md` to cover `VenioNSRLTool` (currently undocumented) and link to `LICENSE`.
- **[CHORE]** Add `CHANGELOG.md` and (optionally) `CONTRIBUTING.md` / issue templates.
- **[ENHANCE]** Publish the PowerShell converter as a module to the PowerShell Gallery once it is cross-platform.
- **[ENHANCE]** Provide sample inputs under `examples/` and an end-to-end smoke test that runs in CI on a Windows runner (with Excel) and a Linux runner (without).
