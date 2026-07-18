# Privacy model

FormaFlow is designed for sensitive school records.

- Processing happens through `127.0.0.1` on the same computer.
- Uploaded bytes are kept only in memory during parsing.
- Parsed sessions are stored only in Python process memory.
- No account, database, cloud storage, telemetry, or external AI is used.
- Restarting the server clears sessions.
- Test fixtures contain synthetic names and values only.

The user remains responsible for access control on the computer and for safely
handling downloaded output files.
