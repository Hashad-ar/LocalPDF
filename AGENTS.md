# LocalPDF repository guide

LocalPDF is a Windows 10/11 x64 utility with exactly three tools: Word to PDF, PDF to Word, and ordered PDF merge. It is offline and must never upload documents or add telemetry, services, startup tasks, or tray processes.

## Architecture

- Native WPF interface and orchestration in `app/`, targeting the Windows-provided .NET Framework 4.8 runtime.
- PDFsharp page-object import for lossless ordered PDF merge.
- A small browser-free Rust helper in `word-helper/` uses raw Windows `IDispatch` for Microsoft Word conversion and owns the new Word PID through a kill-on-close Job Object.
- Per-job work lives only under `%TEMP%\LocalPDF\<job-id>` and is removed on success, failure, cancellation, startup, and shutdown.

## Commands

- Production app: `powershell -File scripts/build-native.ps1 -Configuration Release`
- Tests: set `FrameworkPathOverride` to the restored .NET Framework 4.8 reference folder, run `dotnet build tests/LocalPDF.Tests.csproj -c Release`, then run the resulting test executable.
- Portable build and NSIS installer: `powershell -File scripts/package-native.ps1`

## Non-negotiables

- No network calls, remote assets, telemetry, analytics, OCR, document history, or unrelated PDF tools.
- Never attach to or terminate a user's existing Word process.
- Closing must cancel work and leave no LocalPDF/helper/app-owned Word processes.
- Never overwrite output without approval from the native save dialog and never delete user documents during cleanup.
- Preserve merge order and PDF page objects rather than rendering pages.
