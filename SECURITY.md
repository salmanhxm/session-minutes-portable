# Security Policy

## Supported versions

Only the latest published 1.x release receives security fixes.

## Reporting a vulnerability

Use GitHub Private Vulnerability Reporting for the repository. Do not include real case documents, personal data, credentials, or confidential paths in a public issue. If private reporting is unavailable, open a public issue containing only a request for a private contact channel.

## Security boundaries

- DOCX files and project folders are local, user-selected input.
- Source documents must never be overwritten.
- Generated files must remain under the selected project's `outputs` folder unless the CLI operator explicitly selects another output root.
- The application does not require or store runtime secrets.
- Release executables should be code-signed when a trusted certificate is available; otherwise verify the published SHA-256 checksum.
