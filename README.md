# LocalPDF

LocalPDF is a small, local-first Windows utility for three focused document tasks:

- Convert Microsoft Word documents to PDF.
- Convert digitally generated PDFs to editable Word documents.
- Merge complete PDF pages in an explicitly chosen order.

The application is designed for people who want a straightforward desktop workflow without uploads, accounts, telemetry, browser runtimes, background services, or a permanent local server.

## Key features

- **Word → PDF:** uses the locally installed Microsoft Word rendering engine for high-fidelity output.
- **PDF → Word:** uses Word's native PDF import and DOCX export capabilities.
- **Ordered PDF merge:** supports multiple PDFs, drag-and-drop input, visual reordering, and lossless page-object import.
- **Native Windows interface:** dark WPF UI with native open/save dialogs and Explorer drag-and-drop.
- **Safe cancellation:** conversion and merge work runs in killable helper processes so the UI remains responsive.
- **Output protection:** results are validated in a temporary workspace before an existing destination is atomically replaced.
- **Single-instance behavior:** launching LocalPDF again restores the existing window.
- **No background footprint:** no tray application, service, startup task, updater, or scheduled task remains after the window closes.

## Technology stack

| Layer | Technology |
| --- | --- |
| Desktop UI and orchestration | C#, WPF, .NET Framework 4.8 |
| PDF parsing and merging | PDFsharp WPF 6.2.4 |
| Word automation helper | Rust 1.88, raw Windows `IDispatch` COM automation |
| Process isolation | Windows Job Objects with `JOB_OBJECT_LIMIT_KILL_ON_JOB_CLOSE` |
| Installer | NSIS, current-user installation |
| Tests | Native .NET Framework console test harness |

LocalPDF uses two small native components:

1. `LocalPDF.exe` provides the WPF interface, validation, merge orchestration, temporary storage, and lifecycle management.
2. `LocalPDF-WordHelper.exe` performs Word automation in an isolated process and tracks the exact Word process it creates.

See [ARCHITECTURE.md](ARCHITECTURE.md) for the process, privacy, and output-safety design.

## Repository structure

```text
app/                  WPF application and PDF services
  Assets/             Windows application icon
  Services/           Conversion, merge, process, and temporary-file services
installer/            NSIS installer definition
scripts/              Reproducible build and packaging scripts
tests/                Automated native test harness
word-helper/          Browser-free Rust COM helper for Microsoft Word
ARCHITECTURE.md        Architecture and lifecycle details
IMPLEMENTATION_PLAN.md Implementation and verification outline
PRIVACY.md             Runtime privacy and data-handling statement
SECURITY.md            Security model and reporting guidance
```

Generated files under `bin/`, `obj/`, `target/`, `outputs/`, and `work/` are intentionally excluded from version control.

## System requirements

### Running LocalPDF

- Windows 10 or Windows 11, x64.
- .NET Framework 4.8.
- Microsoft Word desktop application for Word → PDF and PDF → Word.

PDF merging does not require Microsoft Word and remains available when Word is not installed.

### Building from source

- Windows 10 or Windows 11, x64.
- PowerShell.
- .NET SDK 8 or later with NuGet access for the initial restore.
- Rust toolchain 1.88 installed through `rustup`.
- NSIS for installer packaging only.

The first source build downloads declared NuGet and Cargo dependencies. The finished application does not require network access for LocalPDF-owned functionality.

## Installation

Release binaries are intentionally not committed to this repository. Build the installer from source or download a trusted release artifact when one is published.

The NSIS installer:

- installs per user under `%LOCALAPPDATA%\LocalPDF`;
- does not require administrator privileges;
- creates Desktop and Start Menu shortcuts;
- registers a normal Windows uninstaller; and
- removes only LocalPDF-owned application, shortcut, registry, and temporary files.

## Usage

### Word to PDF

1. Open LocalPDF and select **Word → PDF**.
2. Drop a `.doc` or `.docx` file onto the window, or click the drop area to browse.
3. Select **Convert to PDF** and choose the destination.
4. Open the completed file or its containing folder from the result screen.

### PDF to Word

1. Select **PDF → Word**.
2. Choose a digitally generated PDF.
3. Select **Convert to Word** and choose the output path.
4. Review the generated DOCX in Microsoft Word.

LocalPDF warns when a PDF appears to contain pages without selectable text. OCR is not included.

### Merge PDFs

1. Select **Merge PDFs**.
2. Add two or more PDF files.
3. Drag items or use the arrow controls to set the exact page order.
4. Select **Merge PDFs** and choose the output path.

Pages are imported as PDF objects rather than rasterized images, preserving the original page quality.

## Build and development

Clone the repository and enter its directory:

```powershell
git clone https://github.com/Hashad-ar/LocalPDF.git
Set-Location LocalPDF
```

Install the pinned Rust toolchain if needed:

```powershell
rustup toolchain install 1.88.0
```

Build the production application:

```powershell
.\scripts\build-native.ps1 -Configuration Release
```

The application is produced under:

```text
app\bin\Release\net48\
```

The build script places Rust artifacts under `%TEMP%\LocalPDF-build-target`, builds the WPF application, and copies the Word helper beside `LocalPDF.exe`.

### Run the automated tests

Restore the test project and set the .NET Framework reference path used by the WPF build:

```powershell
dotnet restore .\tests\LocalPDF.Tests.csproj
$env:FrameworkPathOverride = Join-Path $env:USERPROFILE '.nuget\packages\microsoft.netframework.referenceassemblies.net48\1.0.3\build\.NETFramework\v4.8'
dotnet build .\tests\LocalPDF.Tests.csproj -c Release --no-incremental
.\tests\bin\Release\net48\LocalPDF.Tests.exe
```

The automated suite covers merge order, Unicode paths, damaged PDFs, scan detection, stale temporary cleanup, helper JSON compatibility, and formatting helpers. Real Word integration should also be smoke-tested on a Windows machine with Word installed.

### Build the portable package and installer

The packaging script currently expects `makensis.exe` at:

```text
%LOCALAPPDATA%\tauri\NSIS\makensis.exe
```

With that compiler available, run:

```powershell
.\scripts\package-native.ps1
```

The portable folder, ZIP archive, and installer are written to the ignored `outputs/` directory.

## Privacy and security

- LocalPDF-owned binaries contain no HTTP client, remote assets, analytics, telemetry, account system, or upload feature.
- Input files are processed locally and are never copied into the repository by the application.
- Per-job working files are stored under `%TEMP%\LocalPDF\<job-id>` and removed after success, failure, or cancellation.
- Stale LocalPDF-owned temporary directories are removed at startup.
- Word documents are opened read-only with alerts and macro automation disabled.
- LocalPDF records Word process IDs before automation and terminates only the new process it created.
- Output is staged and validated before being moved or atomically replaced at the user-approved destination.

Microsoft Word is a separately installed product. Office may make its own licensing or service connections according to the user's Microsoft 365 and Windows policies; LocalPDF does not change system-wide Office or firewall settings. See [PRIVACY.md](PRIVACY.md) and [SECURITY.md](SECURITY.md).

## Known limitations

- Microsoft Word is required for both conversion features.
- PDF merge is available without Word.
- OCR is not included, so scanned PDFs may not become editable text.
- PDF-to-Word quality is determined by Microsoft Word's PDF reflow engine and the structure of the source PDF.
- Password-protected, damaged, or unsupported documents may be rejected.
- The current build and installer are not code-signed and may trigger Windows SmartScreen warnings.
- The packaging script expects NSIS at a fixed local path.
- No release automation or update mechanism is included.

## Potential future improvements

- Add trusted Authenticode signing for the application and installer.
- Make the NSIS compiler path configurable.
- Add a Windows CI workflow for clean build and test verification.
- Publish versioned checksums and signed release artifacts.
- Improve accessibility coverage and add more representative, redistributable PDF fixtures.

OCR, cloud conversion, telemetry, background updates, and unrelated PDF-editing features are intentionally outside the current lightweight scope.

## License

No open-source license has been selected yet. Until a license is added, the source remains under the repository owner's copyright and no redistribution rights are granted by default.
