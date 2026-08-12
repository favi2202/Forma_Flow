from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]


def read(rel: str) -> str:
    return (ROOT / rel).read_text(encoding="utf-8")


def write(rel: str, text: str) -> None:
    (ROOT / rel).write_text(text, encoding="utf-8")


def replace_once(text: str, old: str, new: str, label: str) -> str:
    if new in text:
        return text
    if old not in text:
        raise RuntimeError(f"Could not find expected {label} block")
    return text.replace(old, new, 1)


# Backend: expiring in-memory sessions and explicit deletion.
path = "app.py"
s = read(path)
s = replace_once(s, "import tempfile\n", "import tempfile\nimport time\n", "time import")
s = replace_once(
    s,
    'MAX_FILES = 50\n\napp = FastAPI(title="FormaFlow Local", version="0.6.0")',
    'MAX_FILES = 50\nMAX_SESSIONS = 30\nSESSION_TTL_SECONDS = 30 * 60\n\napp = FastAPI(title="FormaFlow Local", version="0.6.1")',
    "session constants/version",
)
s = replace_once(
    s,
    "SESSIONS: dict[str, dict[str, Any]] = {}\n\nFIELD_DEFINITIONS",
    '''SESSIONS: dict[str, dict[str, Any]] = {}\n\n\ndef cleanup_sessions(now: float | None = None) -> int:\n    """Remove expired in-memory sessions and return the number removed."""\n    current = time.monotonic() if now is None else now\n    expired = [\n        session_id\n        for session_id, session in SESSIONS.items()\n        if current - float(session.get("last_access", session.get("created_at", current))) > SESSION_TTL_SECONDS\n    ]\n    for session_id in expired:\n        SESSIONS.pop(session_id, None)\n    return len(expired)\n\n\ndef get_session(session_id: str) -> dict[str, Any]:\n    """Return a live session and refresh its idle timeout."""\n    now = time.monotonic()\n    cleanup_sessions(now)\n    session = SESSIONS.get(session_id)\n    if session is None:\n        raise HTTPException(status_code=404, detail="Session expired. Upload the files again.")\n    session["last_access"] = now\n    return session\n\n\nFIELD_DEFINITIONS''',
    "session helpers",
)
s = replace_once(
    s,
    '    return {"status": "ok", "version": "0.6.0"}',
    '    cleanup_sessions()\n    return {"status": "ok", "version": "0.6.1"}',
    "health version",
)
s = replace_once(
    s,
    '    SESSIONS[session_id] = {\n        "columns": active["columns"] if active else [],',
    '    now = time.monotonic()\n    cleanup_sessions(now)\n    SESSIONS[session_id] = {\n        "created_at": now,\n        "last_access": now,\n        "columns": active["columns"] if active else [],',
    "session creation",
)
s = replace_once(
    s,
    '''    if len(SESSIONS) > 30:\n        oldest = next(iter(SESSIONS))\n        if oldest != session_id:\n            SESSIONS.pop(oldest, None)''',
    '''    while len(SESSIONS) > MAX_SESSIONS:\n        oldest = min(SESSIONS, key=lambda key: float(SESSIONS[key].get("last_access", 0.0)))\n        if oldest == session_id and len(SESSIONS) == 1:\n            break\n        SESSIONS.pop(oldest, None)''',
    "session cap",
)
old_lookup = '''    session = SESSIONS.get(request.session_id)\n    if session is None:\n        raise HTTPException(status_code=404, detail="Session expired. Upload the files again.")'''
if old_lookup in s:
    s = s.replace(old_lookup, '    session = get_session(request.session_id)')

if '@app.delete("/api/session/{session_id}"' not in s:
    marker = '\n\nif __name__ == "__main__":\n'
    if marker not in s:
        raise RuntimeError("Could not find __main__ block")
    endpoint = '''\n\n@app.delete("/api/session/{session_id}", status_code=204)\ndef delete_session(session_id: str) -> Response:\n    """Explicitly discard a temporary session and its parsed records."""\n    SESSIONS.pop(session_id, None)\n    return Response(status_code=204, headers={"Cache-Control": "no-store"})\n'''
    s = s.replace(marker, endpoint + marker, 1)
write(path, s)

# Tests for explicit deletion and idle expiration.
path = "tests/test_app.py"
s = read(path)
s = replace_once(
    s,
    "from app import app, next_class_value, rows_to_records, is_repeated_header_row",
    '''from app import (\n    SESSION_TTL_SECONDS,\n    SESSIONS,\n    app,\n    cleanup_sessions,\n    is_repeated_header_row,\n    next_class_value,\n    rows_to_records,\n)''',
    "test imports",
)
if "def test_session_can_be_deleted_explicitly()" not in s:
    s += '''\n\ndef test_session_can_be_deleted_explicitly():\n    response = client.post(\n        "/api/upload",\n        files=[("files", ("class8.xlsx", make_xlsx(), "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"))],\n    )\n    assert response.status_code == 200, response.text\n    session_id = response.json()["session_id"]\n    assert session_id in SESSIONS\n\n    deleted = client.delete(f"/api/session/{session_id}")\n    assert deleted.status_code == 204\n    assert deleted.headers["cache-control"] == "no-store"\n    assert session_id not in SESSIONS\n\n    preview = client.post(\n        "/api/preview",\n        json={\n            "session_id": session_id,\n            "columns": [{"key": "student_name", "name": "Student name"}],\n            "fixed_columns": [],\n            "derived_columns": [],\n            "options": {},\n            "limit": 10,\n        },\n    )\n    assert preview.status_code == 404\n\n\ndef test_idle_session_cleanup_expires_old_records():\n    session_id = "expired-test-session"\n    SESSIONS[session_id] = {\n        "created_at": 100.0,\n        "last_access": 100.0,\n        "columns": [],\n        "records": [],\n        "dataset_groups": {},\n    }\n    removed = cleanup_sessions(now=100.0 + SESSION_TTL_SECONDS + 1)\n    assert removed >= 1\n    assert session_id not in SESSIONS\n'''
write(path, s)

# Version and documentation.
write("VERSION", "0.6.1\n")

path = "README.md"
s = read(path)
s = s.replace("favi2202/formaflow", "favi2202/Forma_Flow")
s = s.replace(
    "No cloud AI, accounts, analytics, or external file upload are used.",
    "Local desktop mode uses no cloud AI, accounts, analytics, or external file upload. The Android client uses the configured HTTPS FormaFlow server for processing.",
)
if "sessions expire automatically after 30 minutes of inactivity;" not in s:
    s = s.replace(
        "- sessions exist only in server memory;\n",
        "- sessions exist only in server memory;\n- sessions expire automatically after 30 minutes of inactivity;\n- sessions can be explicitly deleted by web/mobile clients;\n",
    )
if "## Android app" not in s:
    android = '''\n## Android app\n\nThe repository also contains a native Kotlin + Jetpack Compose client under `android/`. It connects to the FastAPI backend, supports Android file picking, native preview, dataset switching, and XLSX/CSV/DOCX export.\n\nAndroid/server mode uploads only the files selected by the user to the configured HTTPS backend. The desktop/local mode remains fully local.\n\n'''
    s = s.replace("## Run on Windows\n", android + "## Run on Windows\n", 1)
write(path, s)

path = "docs/PRIVACY.md"
s = read(path)
if "## Temporary server sessions" not in s:
    s += '''\n\n## Temporary server sessions\n\nWhen FormaFlow is deployed as a server, parsed records are held only in process memory. Sessions expire after 30 minutes of inactivity and can be explicitly deleted with `DELETE /api/session/{session_id}`. Restarting the server clears every session.\n\nThe Android client is an online client: files selected by the user are sent over HTTPS to the configured FormaFlow backend for parsing. Use the local desktop mode when files must never leave the device.\n'''
write(path, s)

path = "CHANGELOG.md"
s = read(path)
if "## v0.6.1" not in s:
    entry = '''## v0.6.1 — Session Privacy & Mobile Hardening\n\n- Added 30-minute idle expiration for in-memory sessions.\n- Added explicit `DELETE /api/session/{session_id}` cleanup.\n- Active sessions refresh their idle timeout on dataset selection, preview, and export.\n- Added automated tests for session expiration and explicit deletion.\n- Documented the Android/server privacy boundary.\n\n'''
    s = s.replace("# Changelog\n\n", "# Changelog\n\n" + entry, 1)
write(path, s)
