# Release Checklist

- [ ] `VERSION` and `CHANGELOG.md` agree.
- [ ] No real DOCX, output, local path, credential, certificate, cache, executable, or ZIP is tracked.
- [ ] PowerShell scripts parse without errors.
- [ ] Python compile and unit tests pass.
- [ ] Source lint/static checks pass or exceptions are documented.
- [ ] Pinned dependencies restore and audit cleanly.
- [ ] `build.ps1` succeeds with CPython 3.12 x64.
- [ ] Launcher and engine show the expected product/file version.
- [ ] Portable smoke tests pass from a separate extracted folder.
- [ ] Preview, optional pilot, direct final generation, source immutability, and Word review are checked.
- [ ] Arabic RTL, mixed paths, 1080p/2K/4K, 100%/125%/150%/250% DPI, and multi-monitor movement are manually checked where hardware is available.
- [ ] ZIP SHA-256 matches the published `.sha256` file.
- [ ] Authenticode signatures are valid, or the unsigned-build limitation is stated.
- [ ] Release notes contain no personal or case data.
