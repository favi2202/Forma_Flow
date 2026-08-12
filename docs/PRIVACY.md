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


## Temporary server sessions

When FormaFlow is deployed as a server, parsed records are held only in process memory. Sessions expire after 30 minutes of inactivity and can be explicitly deleted with `DELETE /api/session/{session_id}`. Restarting the server clears every session.

The Android client is an online client: files selected by the user are sent over HTTPS to the configured FormaFlow backend for parsing. Use the local desktop mode when files must never leave the device.
