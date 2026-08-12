from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
ANDROID = ROOT / "android"


def load(rel: str) -> str:
    return (ANDROID / rel).read_text(encoding="utf-8")


def save(rel: str, text: str) -> None:
    (ANDROID / rel).write_text(text, encoding="utf-8")


def replace_once(text: str, old: str, new: str, label: str) -> str:
    if new in text:
        return text
    if old not in text:
        raise RuntimeError(f"Could not find expected {label} block")
    return text.replace(old, new, 1)


path = "app/src/main/java/uz/formaflow/mobile/data/FormaFlowApi.kt"
s = load(path)
s = s.replace('setRequestProperty("User-Agent", "FormaFlow-Android/0.1")', 'setRequestProperty("User-Agent", "FormaFlow-Android/0.1.1")')
if "fun deleteSession(" not in s:
    marker = '    fun export(baseUrl: String, payload: JSONObject, fallbackFormat: String): ExportResult {\n'
    method = '''    fun deleteSession(baseUrl: String, sessionId: String) {\n        val connection = open(baseUrl, "/api/session/$sessionId", "DELETE")\n        connection.useConnection { code, body, _ ->\n            ensureSuccess(code, body)\n            Unit\n        }\n    }\n\n'''
    s = replace_once(s, marker, method + marker, "deleteSession insertion")
save(path, s)

path = "app/src/main/java/uz/formaflow/mobile/ui/FormaFlowViewModel.kt"
s = load(path)
old = '''    fun clearAll() {\n        _state.update {\n            FormaFlowUiState(baseUrl = it.baseUrl, serverState = it.serverState, serverVersion = it.serverVersion)\n        }\n    }\n'''
new = '''    fun clearAll() {\n        val snapshot = _state.value\n        _state.update {\n            FormaFlowUiState(baseUrl = it.baseUrl, serverState = it.serverState, serverVersion = it.serverVersion)\n        }\n        snapshot.sessionId?.let { sessionId ->\n            viewModelScope.launch(Dispatchers.IO) {\n                runCatching { api.deleteSession(snapshot.baseUrl, sessionId) }\n            }\n        }\n    }\n'''
s = replace_once(s, old, new, "clearAll")
old = '                if (response.columns.isNotEmpty()) refreshPreview()\n'
new = '''                snapshot.sessionId?.takeIf { it != response.sessionId }?.let { oldSessionId ->\n                    viewModelScope.launch(Dispatchers.IO) {\n                        runCatching { api.deleteSession(snapshot.baseUrl, oldSessionId) }\n                    }\n                }\n                if (response.columns.isNotEmpty()) refreshPreview()\n'''
s = replace_once(s, old, new, "old-session cleanup")
save(path, s)

path = "app/src/main/java/uz/formaflow/mobile/MainActivity.kt"
s = load(path)
s = s.replace(
    "Temporary sessions are kept in Render memory. If the server restarts, upload the files again.",
    "Temporary server sessions expire after 30 minutes of inactivity. Clear removes the active session immediately.",
)
save(path, s)

path = "app/build.gradle.kts"
s = load(path)
s = s.replace("versionCode = 1", "versionCode = 2")
s = s.replace('versionName = "0.1.0"', 'versionName = "0.1.1"')
s = s.replace("compileSdk = 37", "compileSdk = 36")
save(path, s)

path = "README.md"
s = load(path)
s = s.replace("# FormaFlow Android v0.1", "# FormaFlow Android v0.1.1")
s = s.replace(
    "- Render sessions are in memory and disappear when the service restarts.",
    "- Render sessions are memory-only, expire after 30 minutes of inactivity, and are cleared when the user taps Clear.",
)
s = s.replace("Allow Gradle sync to install Android SDK 37 and dependencies.", "Allow Gradle sync to install Android SDK 36 and dependencies.")
save(path, s)

path = "CHANGELOG.md"
s = load(path)
if "## 0.1.1" not in s:
    body = s.split("\n", 2)[-1] if s.startswith("# Changelog") else s
    s = '''# Changelog\n\n## 0.1.1 - 2026-08-12\n\n- Added explicit server-session deletion from the Android client.\n- Clear now removes parsed server data instead of only resetting the screen.\n- Re-uploading replaces and discards the previous temporary server session.\n- Updated privacy messaging for the 30-minute backend session timeout.\n- Builds directly from readable repository source.\n\n''' + body.lstrip("\n")
save(path, s)
