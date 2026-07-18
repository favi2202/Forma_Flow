# Changelog

## v0.6.0 — Document Intelligence

- Added deterministic document classification with confidence and reasons.
- Added categories for rosters, promotion documents, monitoring, assessment
  results, meeting minutes, methodical documents, financial documents, lesson
  materials, and unknown documents.
- Added dataset grouping so incompatible table structures are never merged
  blindly.
- Added an interface for switching between detected dataset groups.
- Added multi-sheet workbook processing.
- Added PDF text and table import through `pdfplumber`.
- Added scanned-PDF warnings when no selectable content is available.
- Added DOCX XML recovery when damaged media prevents normal Word parsing.
- Added one-column roster recognition.
- Added headerless promotion-table recognition.
- Added value-based inference for row number, student name, current class, and
  promoted class.
- Fixed numeric monitoring data rows being mistaken for headers.
- Preserved distinct `question_1`, `question_2`, and later monitoring fields.
- Added empty monitoring-template suppression.
- Added stronger filtering of methodical/report text falsely appearing as a
  student roster.
- Added dominant-table document classification while preserving meeting-minute
  and methodical container types.
- Added richer per-file diagnostics, warnings, table counts, source counts, and
  cleanup statistics.
- Added PDF to the English, Uzbek, and Russian interface.
- Increased batch limit to 50 files.
- Added nine v0.6 regression tests; total: 27 passing tests.
- Stress-tested against a 33-file mixed local corpus without packaging private
  source documents.

## v0.5.0 — Smart Import

- Fixed Word rosters without a table-header row.
- Preserved the first student in headerless Word lists.
- Removed columns empty across an entire table.
- Added row-number recognition for `№`, `T/N`, and related labels.
- Inferred missing class values from titles, filenames, and sheet names.
- Rebuilt HTML-as-XLS parsing with `rowspan` and `colspan` support.
- Preserved guardian phone and email columns.
- Merged continuation contact rows belonging to one student.
- Marked sensitive fields and left them unchecked by default.

## v0.4.0

- Added DOCX table import and optional legacy DOC conversion through LibreOffice.
- Detected, merged, and flagged multiple Word tables.

## v0.3.1

- Removed repeated main headings and secondary labels from exported records.

## v0.3.0

- Added English, Uzbek, and Russian interfaces.
- Added next-class and other calculated columns.
- Added cleanup, sorting, preview, and export improvements.
