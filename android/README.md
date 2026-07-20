# FormaFlow Android v0.1

A native Android client for the existing FormaFlow FastAPI/Render backend.

## What works

- Native Kotlin + Jetpack Compose interface (not a WebView)
- HTTPS server health check, including Render cold-start messaging
- Android multi-file picker for XLSX/XLS/XLSM/CSV/DOCX/DOC/PDF
- Multipart upload progress
- Server-side document classification and table extraction
- Dataset switching
- Select and rename output columns
- Trim-whitespace and remove-duplicate options
- Native preview table
- XLSX, CSV, and DOCX export through Android's save-file picker
- Editable backend URL saved locally
- No app analytics, accounts, or advertising

## Default backend

`https://forma-flow.onrender.com/`

The URL can be changed from inside the app. It must use HTTPS.

## Privacy

This build uses **online processing**. Selected documents are uploaded to the configured FormaFlow server. Do not use real sensitive student data until the server deployment, retention behavior, access controls, and privacy policy have been reviewed.

The app itself does not make private copies of imported documents. The Android file picker grants read access only to files selected by the user. Export bytes are held in memory until the user chooses a save destination.

## Open and build

1. Install the latest stable Android Studio.
2. Open this folder as an Android Studio project.
3. Allow Gradle sync to install Android SDK 37 and dependencies.
4. Connect an Android phone with USB debugging or start an emulator.
5. Run the `app` configuration.

Debug APK:

```text
app/build/outputs/apk/debug/app-debug.apk
```

Command line:

```bash
./gradlew assembleDebug
```

## Requirements

- Android 8.0 (API 26) or newer
- Internet connection
- Working FormaFlow v0.6 backend

## Current limitations

- The free Render instance may sleep and make the first request slow.
- Render sessions are in memory and disappear when the service restarts.
- Processing speed is still limited by the Render server.
- Fixed columns, derived columns, custom sorting, and multilingual Android UI are planned for later versions.
- This is a development/debug build until a private release signing key is configured.

## Repository layout

- `app/src/main/java/.../data/` — API and models
- `app/src/main/java/.../ui/` — state and Compose theme
- `MainActivity.kt` — native mobile workflow and screens

## Repository location

This Android project lives in the `android/` directory of the FormaFlow repository. The backend remains at the repository root and is deployed separately on Render.
