# Contributing

1. Use Windows 11 x64 and CPython 3.12 x64.
2. Never commit real DOCX files, case identifiers, generated outputs, local paths, credentials, certificates, or build caches.
3. Keep business behavior, CLI arguments, JSON schemas, and OOXML preservation rules backward compatible unless the change is explicitly documented.
4. Run the PowerShell parser, Python tests, source-quality checks, and a portable build before opening a pull request.
5. Use synthetic documents only in tests.

See `BUILD_WINDOWS.md` for exact commands. `SESSION_MINUTES_UI_DIAGNOSTIC=1` is an optional local diagnostic flag; it is not required at runtime.
