# Import engine overview

FormaFlow follows this pipeline:

```text
file signature inspection
→ format-specific extraction
→ document classification
→ table inventory
→ header and schema detection
→ column normalization
→ row repair and cleanup
→ dataset classification
→ compatible dataset grouping
→ user review
→ preview/export
```

## Safety principles

1. File contents, not only extensions, determine the parser.
2. Unknown fields are preserved instead of silently deleted.
3. Incompatible datasets stay separate.
4. Sensitive fields are not selected by default.
5. Source files are never modified.
6. Automatic inference remains visible and reviewable.

## Adding a canonical field

1. Add the field and aliases to `FIELD_DEFINITIONS` in `app.py`.
2. Add its translated label to `static/app.js`.
3. Decide whether it is core, sensitive, or neither in `merge_parsed_files`.
4. Add synthetic regression tests.

## Classification

Document classification combines filename/text keywords and the distribution of
recognized table types. Strong narrative containers such as meeting minutes and
methodical documents remain the document type even when they contain large data
appendices.

Dataset classification uses canonical field combinations. For example:

- student name + class → roster;
- current class + promoted class → promotion table;
- student name + question/result columns → monitoring;
- financial vocabulary → financial table.
