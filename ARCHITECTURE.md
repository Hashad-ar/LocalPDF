# Architecture

## Native privacy boundary

LocalPDF is a native x64 WPF application targeting the Windows-provided .NET Framework 4.8 runtime. It contains no browser, WebView, HTTP client, updater, telemetry SDK, account system, service, tray process, or scheduled task. For LocalPDF-owned binaries, the zero-network guarantee is structural rather than dependent on runtime switches or machine-wide policy.

Microsoft Word is an independently installed dependency. LocalPDF does not request online features or upload documents, but Office can make its own required licensing/service connections according to Microsoft and the machine's Office policy. LocalPDF deliberately does not alter system-wide Office or firewall policy.

The production build contains the normal WPF executable and a small browser-free Rust Word helper. PDF merge work uses the main executable's `--merge-helper` mode; Word conversion is delegated to `LocalPDF-WordHelper.exe`. Heavy work therefore runs outside the UI process so Cancel and window close remain deterministic.

## Conversion and process ownership

Word conversions use local Microsoft Word automation only. Before Word starts, LocalPDF snapshots all existing `WINWORD.EXE` process IDs. The newly created Word process is placed in a Windows Job Object configured with `JOB_OBJECT_LIMIT_KILL_ON_JOB_CLOSE`; its PID is also reported to the parent as a narrow fallback. Existing Word processes are never terminated. If Word does not create an isolated process, the operation stops safely.

The helper opens documents read-only with alerts and macro automation disabled. Word-to-PDF uses `ExportAsFixedFormat`; PDF-to-Word uses Word's native PDF import and DOCX save path. Raw `IDispatch` objects are released deterministically.

## PDF integrity and output safety

PDF merging uses PDFsharp import mode and copies complete page object trees in the displayed order. Pages are not rendered or recompressed. PDF signatures and parsability are validated before and after work.

Every operation writes first to `%TEMP%\LocalPDF\<random-id>`. Only a validated, non-empty result is copied to a same-directory pending file beside the chosen destination, then moved or atomically replaced. Cancellation and failures therefore do not leave partial output. Workspaces are deleted on completion and stale app-owned workspaces are purged at startup.

## Lifecycle

A named mutex enforces one UI instance; a second launch restores the existing window. Closing the window cancels the helper, which closes its Word Job Object, and removes the temporary workspace. The application installs no background components and runs nothing after its window closes.
