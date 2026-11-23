# LocalPDF privacy statement

LocalPDF is designed around local document processing and data minimization.

## What LocalPDF collects

Nothing. LocalPDF has no account system, analytics, telemetry, advertising, crash-reporting service, document history, or remote logging endpoint.

## Network behavior

LocalPDF-owned executables do not include HTTP clients, remote assets, an updater, or an upload path. PDF merging is performed entirely inside the local application.

Word conversion uses the Microsoft Word desktop application installed on the computer. Microsoft Word is an independent product and may contact Microsoft for licensing or required services under the user's Office configuration. LocalPDF does not upload the selected document, request Office connected experiences, or modify global Office and firewall policies.

Users requiring a fully air-gapped conversion environment should disconnect the machine from the network or enforce an organization-managed network policy for Microsoft Word.

## File handling

- Source documents remain at their original paths.
- Temporary job data is written only under `%TEMP%\LocalPDF\<random-job-id>`.
- Temporary job directories are removed after success, failure, or cancellation.
- Stale LocalPDF-owned temporary directories are removed at startup.
- Completed output is written only to the path selected in the native Windows save dialog.
- LocalPDF does not maintain a recent-document database.

## Microsoft Word process handling

LocalPDF snapshots existing `WINWORD.EXE` process IDs before starting automation. It assigns only the newly created Word process to a kill-on-close Windows Job Object. Existing user-owned Word processes and documents are not attached to or terminated.

## Repository hygiene

Build outputs, local workspaces, test documents, diagnostic data, environment files, credentials, signing keys, and IDE-specific state are excluded by `.gitignore`. Real personal documents should never be committed as test fixtures.
