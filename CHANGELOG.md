# Changelog

## 1.2.5 - 2026-07-13

- Made the verified responsive WinForms interface a first-class source file so clean builds no longer depend on a historical `dist` folder.
- Prevented the Arabic review heading from clipping by measuring its row from the active font and DPI.
- Imported the signature table's complete Word style closure and fixed its grid to the template's physical column widths.
- Removed the template's absolute page-Y table anchor so signatures flow immediately after the closing paragraph without a blank-page gap.
- Preserved all closing text, four signature images, role labels, member names, dropdown definitions, and package relationships.
- Added regression coverage for user-edited dropdown lists, style conflicts, fixed signature geometry, and inline centered flow.

## 1.2.2 - 2026-07-13

- Packaged the complete, visually verified 0.6.8 WinForms implementation as the UI-only baseline while preserving the current launcher and document engine.
- Preserved the daily template's floating signature-table geometry, row and cell properties, font sizes, image sizes, and positioning instead of converting it to inline flow.
- Replaced an older matching signature block with a fresh exact clone of the current daily template and remapped only package relationships and volatile drawing identifiers.
- Added structural validation requiring every generated signature table to match the current template layout and image geometry.
- Verified a 70-document isolated final run with zero blocked outputs.

## 1.2.1 - 2026-07-13

- Restored the proven 0.6.8 WinForms DPI and window-capacity profile to prevent overlapping or oversized text in normal and maximized windows.
- Preserved the modern workflow, status grid, portable launcher path handling, and current processing features while reverting only the sizing system.
- Prevented long Arabic committee-member names from losing their final glyphs by applying a conservative one-point reduction only to long signature-name cells.
- Added regression coverage proving user-edited ComboBox and DropDownList options are discovered from the daily template and preserved in generated documents.

## 1.2.0 - 2026-07-12

- Reworked WinForms scaling around logical screen capacity and monitor DPI so typography, cards, icons, the activity grid, and the footer remain proportional when the window or display scale changes.
- Reduced compact-window density while preserving complete Arabic RTL labels and isolated LTR paths.
- Copied the complete closing signature table and its four images immediately after the closing paragraph while preserving its native OOXML formatting.
- Normalized floating signature tables into the document flow, prevented duplicate closing blocks, and made repeated runs idempotent.
- Allocated drawing identifiers across the complete DOCX package and validated imported relationships, media, content types, and identifiers before saving.
- Raised the preview schema to version 6 so approval fingerprints include the final closing-block behavior.

## 1.1.1 - 2026-07-12

- Measured activity-grid header height from the active font and monitor DPI.
- Reserved a non-wrapping footer area for the complete Made By Salman credit.

## 1.1.0 - 2026-07-12

- Added complete closing-block transfer: the decision paragraph, floating signature table, four images, five relationships, and required content types are now copied together.
- Added package validation for exactly one signature table and every imported image part while preserving source files.
- Replaced fixed-height assumptions with DPI-aware preferred-text measurements and automatic scrolling when the available screen cannot contain the full workflow.
- Raised the preview schema to version 5 so approvals include the signature-table fingerprint and relationship count.

## 1.0.6 - 2026-07-11

- Guaranteed enough responsive width for the complete Arabic application name in the top navigation.

## 1.0.5 - 2026-07-11

- Expanded the top brand area so the full Arabic application name remains visible.
- Reserved complete rows for the review guidance and bottom status text in normal and maximized windows.

## 1.0.4 - 2026-07-11

- Switched Arabic labels to GDI+ compatible text rendering to preserve complete glyph bounds at fractional DPI.
- Removed automatic ellipsis from primary interface labels so text is never silently cropped.

## 1.0.3 - 2026-07-11

- Added Arabic-glyph headroom to navigation, hero, workflow-card, and activity-log title rows.
- Tuned compact-window typography independently from container scaling to avoid overlap and truncation.

## 1.0.2 - 2026-07-11

- Calculated responsiveness from the logical client area and applied monitor DPI only to pixel dimensions.
- Prevented Arabic title, card, and activity-log text clipping on high-DPI normal and maximized windows.

## 1.0.1 — 2026-07-11

- Fixed double DPI scaling that enlarged and clipped Arabic text in both the normal and maximized window.
- Kept responsive row, padding, and font calculations under one DPI-aware layout system.

## 1.0.0 — 2026-07-11

- Prepared the first audited portable Windows release.
- Added a standalone project layout, local assets, release metadata, x64/CPython checks, and SHA-256 output.
- Updated Arabic and English workflow documentation: preview is required, pilot is optional, and human Word review remains required.
- Added GitHub-ready ignore rules, CI, contribution, security, asset, build, and release documentation.
- Updated `lxml` to 6.1.0.
- Preserved all existing OOXML processing and source-protection behavior.
