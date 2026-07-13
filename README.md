# Session Minutes Portable

An Arabic RTL Windows 11 desktop application that creates review copies of Word session minutes while preserving OOXML content controls and formatting. Source documents are never overwritten; generated files stay under the selected project's `outputs` folder.

## Runtime requirements

- Windows 11 x64
- Microsoft Word for visual and human review
- No Python installation, administrator rights, cloud account, or internet connection at runtime

## Workflow

1. Extract the complete release archive.
2. Run `SessionMinutesPortable.exe` and select the project folder.
3. Run Preview and inspect the report.
4. Optionally create a pilot set.
5. Create final copies only after validating the preview.
6. Review every generated DOCX in Word.

The project folder contains one numeric batch folder (for example `17`) with numbered subfolders (`1`, `2`, `3`, ...), plus the filling template in the project root. Template opening blocks map to numbered subfolders in document order.

## Privacy and safety

Processing is local. The application does not upload documents, overwrite source DOCX files, or require secrets. Real case files and generated outputs are intentionally excluded from this repository.

## Build

See [BUILD_WINDOWS.md](BUILD_WINDOWS.md). Release builds require CPython 3.12 x64 and produce a ZIP plus a SHA-256 checksum.

## License

MIT. See [LICENSE](LICENSE) and [THIRD_PARTY_NOTICES.md](THIRD_PARTY_NOTICES.md).
