# Security and privacy

FormaFlow processes documents locally. Please do not report security issues by
attaching real school records to a public GitHub issue.

For a potential vulnerability, open a minimal report that contains no private
student or staff data. Use synthetic files when a reproduction document is
needed.

## Data-handling expectations

- The application binds to `127.0.0.1` by default.
- Uploaded records remain in process memory for the current session.
- Restarting the application clears session data.
- No analytics or cloud AI APIs are used.
- Users should review automatic classifications and mappings before exporting
  official records.
