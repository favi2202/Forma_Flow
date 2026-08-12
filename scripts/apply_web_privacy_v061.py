from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]


def load(rel: str) -> str:
    return (ROOT / rel).read_text(encoding="utf-8")


def save(rel: str, text: str) -> None:
    (ROOT / rel).write_text(text, encoding="utf-8")


def replace_once(text: str, old: str, new: str, label: str) -> str:
    if new in text:
        return text
    if old not in text:
        raise RuntimeError(f"Could not find expected {label} block")
    return text.replace(old, new, 1)


# Backend response hardening and aggregate upload guard.
path = "app.py"
s = load(path)
s = replace_once(
    s,
    "MAX_FILE_BYTES = 25 * 1024 * 1024\nMAX_ROWS_PER_FILE = 50_000",
    "MAX_FILE_BYTES = 25 * 1024 * 1024\nMAX_TOTAL_UPLOAD_BYTES = 75 * 1024 * 1024\nMAX_ROWS_PER_FILE = 50_000",
    "aggregate upload limit",
)
if '@app.middleware("http")' not in s:
    marker = 'app.mount("/static", StaticFiles(directory=STATIC_DIR), name="static")\n'
    middleware = '''app.mount("/static", StaticFiles(directory=STATIC_DIR), name="static")\n\n\n@app.middleware("http")\nasync def privacy_security_headers(request: Any, call_next: Any) -> Response:\n    response = await call_next(request)\n    if request.url.path.startswith("/api/") or request.url.path == "/health":\n        response.headers["Cache-Control"] = "no-store"\n    response.headers["X-Content-Type-Options"] = "nosniff"\n    response.headers["Referrer-Policy"] = "no-referrer"\n    response.headers["X-Frame-Options"] = "DENY"\n    return response\n'''
    s = replace_once(s, marker, middleware, "security middleware")

s = replace_once(
    s,
    "    parsed_count = 0\n\n    for upload_file in files:",
    "    parsed_count = 0\n    total_upload_bytes = 0\n\n    for upload_file in files:",
    "upload byte counter",
)
s = replace_once(
    s,
    '''        content = await upload_file.read()\n        if len(content) > MAX_FILE_BYTES:''',
    '''        content = await upload_file.read()\n        total_upload_bytes += len(content)\n        if total_upload_bytes > MAX_TOTAL_UPLOAD_BYTES:\n            raise HTTPException(status_code=413, detail="Combined upload exceeds 75 MB.")\n        if len(content) > MAX_FILE_BYTES:''',
    "combined upload check",
)
save(path, s)

# Web client: tell the truth about local vs hosted processing and clean sessions.
path = "static/app.js"
s = load(path)
if "const isLocalRuntime" not in s:
    s = s.replace(
        "const state = {",
        'const isLocalRuntime = ["localhost", "127.0.0.1", "::1"].includes(window.location.hostname);\n\nconst state = {',
        1,
    )

replacements = {
    'version: "Local Document Intelligence v0.6.0",': 'version: "Document Intelligence v0.6.1",',
    'version: "Mahalliy Document Intelligence v0.6.0",': 'version: "Document Intelligence v0.6.1",',
    'version: "Локальный Document Intelligence v0.6.0",': 'version: "Document Intelligence v0.6.1",',
    'privacyChip: "● Files stay on this computer",': 'privacyChip: "● Files stay on this computer",\n    privacyHosted: "● Files are processed on this server",',
    'privacyChip: "● Fayllar shu kompyuterda qoladi",': 'privacyChip: "● Fayllar shu kompyuterda qoladi",\n    privacyHosted: "● Fayllar ushbu serverda qayta ishlanadi",',
    'privacyChip: "● Файлы остаются на этом компьютере",': 'privacyChip: "● Файлы остаются на этом компьютере",\n    privacyHosted: "● Файлы обрабатываются на этом сервере",',
    'worksLocally: "Works locally",': 'worksLocally: "Works locally",\n    worksHosted: "Server-assisted",',
    'worksLocally: "Mahalliy ishlaydi",': 'worksLocally: "Mahalliy ishlaydi",\n    worksHosted: "Server orqali ishlaydi",',
    'worksLocally: "Работает локально",': 'worksLocally: "Работает локально",\n    worksHosted: "Обработка на сервере",',
    'serverConnected: "Local server connected",': 'serverConnected: "Local server connected",\n    serverConnectedHosted: "Hosted server connected",',
    'serverConnected: "Mahalliy server ulandi",': 'serverConnected: "Mahalliy server ulandi",\n    serverConnectedHosted: "Server ulandi",',
    'serverConnected: "Локальный сервер подключен",': 'serverConnected: "Локальный сервер подключен",\n    serverConnectedHosted: "Сервер подключен",',
    'serverHelp: "Start the app with run.bat or python app.py, then open http://127.0.0.1:8000",': 'serverHelp: "Start the app with run.bat or python app.py, then open http://127.0.0.1:8000",\n    serverHelpHosted: "The hosted service is unavailable. Try again shortly.",',
    'serverHelp: "run.bat yoki python app.py ni ishga tushiring, keyin http://127.0.0.1:8000 manzilini oching",': 'serverHelp: "run.bat yoki python app.py ni ishga tushiring, keyin http://127.0.0.1:8000 manzilini oching",\n    serverHelpHosted: "Server vaqtincha ishlamayapti. Birozdan keyin qayta urinib ko‘ring.",',
    'serverHelp: "Запустите run.bat или python app.py и откройте http://127.0.0.1:8000",': 'serverHelp: "Запустите run.bat или python app.py и откройте http://127.0.0.1:8000",\n    serverHelpHosted: "Сервис временно недоступен. Попробуйте ещё раз чуть позже.",',
    'cleared: "Local page data cleared",': 'cleared: "Local page data cleared",\n    clearedHosted: "Temporary server data cleared",',
    'cleared: "Mahalliy sahifa ma’lumotlari tozalandi",': 'cleared: "Mahalliy sahifa ma’lumotlari tozalandi",\n    clearedHosted: "Serverdagi vaqtinchalik ma’lumotlar tozalandi",',
    'cleared: "Локальные данные страницы очищены",': 'cleared: "Локальные данные страницы очищены",\n    clearedHosted: "Временные данные на сервере очищены",',
}
for old, new in replacements.items():
    if new not in s:
        if old not in s:
            raise RuntimeError(f"Missing frontend text: {old}")
        s = s.replace(old, new, 1)

old_apply = '''  $("languageSelect").value = state.language;\n  const defaultAcademic = fixedList.querySelector('[data-default-key="academicYear"]');'''
new_apply = '''  $("languageSelect").value = state.language;\n  const privacyChip = document.querySelector('[data-i18n="privacyChip"]');\n  if (privacyChip) privacyChip.textContent = isLocalRuntime ? t("privacyChip") : t("privacyHosted");\n  const runtimeMode = document.querySelector('[data-i18n="worksLocally"]');\n  if (runtimeMode) runtimeMode.textContent = isLocalRuntime ? t("worksLocally") : t("worksHosted");\n  const defaultAcademic = fixedList.querySelector('[data-default-key="academicYear"]');'''
s = replace_once(s, old_apply, new_apply, "runtime privacy labels")

s = replace_once(
    s,
    '    status.textContent = t("serverConnected");',
    '    status.textContent = isLocalRuntime ? t("serverConnected") : t("serverConnectedHosted");',
    "connected server label",
)
s = replace_once(
    s,
    '    showMessage(t("serverHelp"), "error");',
    '    showMessage(isLocalRuntime ? t("serverHelp") : t("serverHelpHosted"), "error");',
    "hosted server help",
)

# Delete the previous server session after a successful replacement upload.
old_upload = '''    state.sessionId = payload.session_id;\n    state.files = payload.files || [];'''
new_upload = '''    const previousSessionId = state.sessionId;\n    state.sessionId = payload.session_id;\n    if (previousSessionId && previousSessionId !== state.sessionId) {\n      fetch(`/api/session/${encodeURIComponent(previousSessionId)}`, { method: "DELETE" }).catch(() => {});\n    }\n    state.files = payload.files || [];'''
s = replace_once(s, old_upload, new_upload, "replacement session cleanup")

old_clear = '''$("clearButton").addEventListener("click", () => {\n  state.sessionId = null;'''
new_clear = '''$("clearButton").addEventListener("click", () => {\n  const sessionToDelete = state.sessionId;\n  state.sessionId = null;'''
s = replace_once(s, old_clear, new_clear, "clear session capture")
s = replace_once(
    s,
    '''  hideMessage();\n  showToast(t("cleared"));\n});''',
    '''  hideMessage();\n  if (sessionToDelete) {\n    fetch(`/api/session/${encodeURIComponent(sessionToDelete)}`, { method: "DELETE" }).catch(() => {});\n  }\n  showToast(isLocalRuntime ? t("cleared") : t("clearedHosted"));\n});''',
    "clear server session",
)
save(path, s)

# Tests for no-store/security headers and aggregate limit constant presence.
path = "tests/test_app.py"
s = load(path)
if "def test_api_responses_are_not_cacheable()" not in s:
    s += '''\n\ndef test_api_responses_are_not_cacheable():\n    response = client.get("/health")\n    assert response.status_code == 200\n    assert response.headers["cache-control"] == "no-store"\n    assert response.headers["x-content-type-options"] == "nosniff"\n    assert response.headers["referrer-policy"] == "no-referrer"\n    assert response.headers["x-frame-options"] == "DENY"\n'''
save(path, s)

path = "CHANGELOG.md"
s = load(path)
needle = "- Documented the Android/server privacy boundary.\n"
extra = "- Hosted web UI now clearly distinguishes server processing from local mode.\n- Web and Android clients explicitly delete replaced/cleared temporary sessions.\n- API responses send no-store and basic browser security headers.\n- Added a 75 MB combined-upload guard for public deployments.\n"
if extra not in s:
    s = s.replace(needle, needle + extra, 1)
save(path, s)
