# Security policy

## Supported version

The current `main` branch is the supported development version. No long-term support or automatic update channel is currently provided.

## Reporting a vulnerability

Report suspected vulnerabilities privately to the repository owner through GitHub. If GitHub private vulnerability reporting is enabled for the repository, prefer that channel.

Do not attach confidential documents, credentials, API keys, or personal data to a report. Use a minimal synthetic fixture when reproduction requires a document.

## Security model

LocalPDF uses the following controls:

- no application-owned network client or remote execution path;
- read-only Word document opening with alerts and macro automation disabled;
- exact before/after Word PID tracking;
- kill-on-close Windows Job Objects for app-owned Word automation;
- helper-process isolation for cancellable long-running work;
- PDF signature and parser validation;
- per-operation randomized temporary directories;
- validation before destination commit; and
- no service, tray process, scheduled task, startup item, or background updater.

## Trust boundaries

- Microsoft Word, .NET Framework, PDFsharp, Rust crates, Windows, and NSIS are external dependencies.
- Microsoft Word may use its own Microsoft 365 licensing or service network connections.
- The current application and installer are not Authenticode-signed.
- Malicious or malformed documents remain untrusted input and should be handled on a maintained Windows installation with current Microsoft Office security updates.

## Publication checklist

Before publishing a commit or release:

1. Run the automated tests and applicable Word integration smoke tests.
2. Review `git status` and the staged diff.
3. Scan staged files for credentials, private keys, tokens, user-specific paths, and personal documents.
4. Confirm `work/`, `outputs/`, `bin/`, `obj/`, and `target/` remain ignored.
5. Publish release binaries only from a clean, reviewed source revision.
