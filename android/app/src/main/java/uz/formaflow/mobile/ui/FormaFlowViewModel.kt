package uz.formaflow.mobile.ui

import android.app.Application
import android.content.Context
import android.database.Cursor
import android.net.Uri
import android.provider.OpenableColumns
import androidx.lifecycle.AndroidViewModel
import androidx.lifecycle.viewModelScope
import kotlinx.coroutines.Dispatchers
import kotlinx.coroutines.flow.MutableStateFlow
import kotlinx.coroutines.flow.StateFlow
import kotlinx.coroutines.flow.asStateFlow
import kotlinx.coroutines.flow.update
import kotlinx.coroutines.launch
import kotlinx.coroutines.withContext
import org.json.JSONArray
import org.json.JSONObject
import uz.formaflow.mobile.BuildConfig
import uz.formaflow.mobile.data.ApiException
import uz.formaflow.mobile.data.ColumnInfo
import uz.formaflow.mobile.data.DatasetGroup
import uz.formaflow.mobile.data.ExportResult
import uz.formaflow.mobile.data.FileSummary
import uz.formaflow.mobile.data.FormaFlowApi
import uz.formaflow.mobile.data.PickedDocument
import uz.formaflow.mobile.data.PreviewResponse


enum class ServerState { UNKNOWN, CHECKING, WAKING, READY, OFFLINE }
enum class WorkPhase { IDLE, UPLOADING, ANALYZING, PREVIEWING, EXPORTING }

data class FormaFlowUiState(
    val baseUrl: String = BuildConfig.FORMAFLOW_BASE_URL,
    val serverState: ServerState = ServerState.UNKNOWN,
    val serverVersion: String = "",
    val documents: List<PickedDocument> = emptyList(),
    val fileSummaries: List<FileSummary> = emptyList(),
    val sessionId: String? = null,
    val datasetGroups: List<DatasetGroup> = emptyList(),
    val activeDatasetId: String? = null,
    val columns: List<ColumnInfo> = emptyList(),
    val rowCount: Int = 0,
    val preview: PreviewResponse? = null,
    val trimWhitespace: Boolean = true,
    val removeDuplicates: Boolean = false,
    val phase: WorkPhase = WorkPhase.IDLE,
    val progress: Float = 0f,
    val statusMessage: String = "",
    val errorMessage: String? = null,
    val pendingExport: ExportResult? = null,
)

class FormaFlowViewModel(application: Application) : AndroidViewModel(application) {
    private val api = FormaFlowApi()
    private val preferences = application.getSharedPreferences("formaflow", Context.MODE_PRIVATE)
    private val _state = MutableStateFlow(
        FormaFlowUiState(baseUrl = preferences.getString("base_url", BuildConfig.FORMAFLOW_BASE_URL) ?: BuildConfig.FORMAFLOW_BASE_URL)
    )
    val state: StateFlow<FormaFlowUiState> = _state.asStateFlow()

    init {
        checkServer()
    }

    fun setBaseUrl(value: String) {
        _state.update { it.copy(baseUrl = value, serverState = ServerState.UNKNOWN, errorMessage = null) }
    }

    fun saveBaseUrl() {
        runCatching { FormaFlowApi.normalizeBaseUrl(_state.value.baseUrl) }
            .onSuccess { normalized ->
                preferences.edit().putString("base_url", normalized).apply()
                _state.update { it.copy(baseUrl = normalized, statusMessage = "Server URL saved.", errorMessage = null) }
                checkServer()
            }
            .onFailure { error(it.message ?: "Invalid server URL.") }
    }

    fun checkServer() {
        if (_state.value.phase != WorkPhase.IDLE) return
        viewModelScope.launch {
            _state.update { it.copy(serverState = ServerState.CHECKING, statusMessage = "Checking server…", errorMessage = null) }
            runCatching {
                withContext(Dispatchers.IO) { api.health(_state.value.baseUrl) }
            }.onSuccess { health ->
                _state.update {
                    it.copy(serverState = ServerState.READY, serverVersion = health.version, statusMessage = "Server ready.")
                }
            }.onFailure { throwable ->
                _state.update {
                    it.copy(
                        serverState = ServerState.OFFLINE,
                        statusMessage = "Server unavailable.",
                        errorMessage = friendlyError(throwable),
                    )
                }
            }
        }
    }

    fun addDocuments(context: Context, uris: List<Uri>) {
        val supported = setOf("xlsx", "xlsm", "xls", "csv", "docx", "doc", "pdf")
        val documents = uris.distinct().mapNotNull { uri ->
            val meta = queryDocument(context, uri)
            val extension = meta.first.substringAfterLast('.', "").lowercase()
            if (extension !in supported) null else PickedDocument(
                uri = uri,
                name = meta.first,
                sizeBytes = meta.second,
                mimeType = context.contentResolver.getType(uri) ?: "application/octet-stream",
            )
        }
        if (documents.isEmpty() && uris.isNotEmpty()) {
            error("Choose Excel, CSV, Word, or PDF files.")
            return
        }
        _state.update {
            it.copy(
                documents = (it.documents + documents).distinctBy { document -> document.uri },
                errorMessage = null,
                statusMessage = "${documents.size} file(s) added.",
            )
        }
    }

    fun clearAll() {
        val snapshot = _state.value
        _state.update {
            FormaFlowUiState(baseUrl = it.baseUrl, serverState = it.serverState, serverVersion = it.serverVersion)
        }
        snapshot.sessionId?.let { sessionId ->
            viewModelScope.launch(Dispatchers.IO) {
                runCatching { api.deleteSession(snapshot.baseUrl, sessionId) }
            }
        }
    }

    fun uploadAndAnalyze(context: Context) {
        val snapshot = _state.value
        if (snapshot.documents.isEmpty()) {
            error("Choose at least one file.")
            return
        }
        if (snapshot.documents.any { it.sizeBytes > 25L * 1024L * 1024L }) {
            error("Each file must be 25 MB or smaller.")
            return
        }
        viewModelScope.launch {
            _state.update {
                it.copy(
                    phase = WorkPhase.UPLOADING,
                    progress = 0f,
                    statusMessage = "Uploading files…",
                    errorMessage = null,
                    serverState = ServerState.WAKING,
                )
            }
            runCatching {
                withContext(Dispatchers.IO) {
                    api.upload(context, snapshot.baseUrl, snapshot.documents) { progress ->
                        _state.update { current ->
                            current.copy(
                                progress = progress,
                                phase = if (progress >= 0.999f) WorkPhase.ANALYZING else WorkPhase.UPLOADING,
                                statusMessage = if (progress >= 0.999f) "Analyzing documents…" else "Uploading files…",
                            )
                        }
                    }
                }
            }.onSuccess { response ->
                _state.update {
                    it.copy(
                        serverState = ServerState.READY,
                        phase = WorkPhase.IDLE,
                        progress = 1f,
                        sessionId = response.sessionId,
                        fileSummaries = response.files,
                        datasetGroups = response.datasetGroups,
                        activeDatasetId = response.activeDatasetId,
                        columns = response.columns,
                        rowCount = response.rowCount,
                        preview = null,
                        statusMessage = response.message ?: "Analysis complete: ${response.rowCount} rows.",
                        errorMessage = null,
                    )
                }
                snapshot.sessionId?.takeIf { it != response.sessionId }?.let { oldSessionId ->
                    viewModelScope.launch(Dispatchers.IO) {
                        runCatching { api.deleteSession(snapshot.baseUrl, oldSessionId) }
                    }
                }
                if (response.columns.isNotEmpty()) refreshPreview()
            }.onFailure { throwable ->
                _state.update {
                    it.copy(phase = WorkPhase.IDLE, progress = 0f, serverState = ServerState.OFFLINE, errorMessage = friendlyError(throwable))
                }
            }
        }
    }

    fun selectDataset(datasetId: String) {
        val snapshot = _state.value
        val sessionId = snapshot.sessionId ?: return
        if (datasetId == snapshot.activeDatasetId) return
        viewModelScope.launch {
            _state.update { it.copy(phase = WorkPhase.PREVIEWING, statusMessage = "Switching dataset…", errorMessage = null) }
            runCatching {
                withContext(Dispatchers.IO) { api.selectDataset(snapshot.baseUrl, sessionId, datasetId) }
            }.onSuccess { response ->
                _state.update {
                    it.copy(
                        phase = WorkPhase.IDLE,
                        activeDatasetId = response.activeDatasetId,
                        columns = response.columns,
                        rowCount = response.rowCount,
                        preview = null,
                        statusMessage = "Dataset selected.",
                    )
                }
                refreshPreview()
            }.onFailure { throwable ->
                _state.update { it.copy(phase = WorkPhase.IDLE, errorMessage = friendlyError(throwable)) }
            }
        }
    }

    fun toggleColumn(key: String, selected: Boolean) {
        _state.update { state ->
            state.copy(columns = state.columns.map { if (it.key == key) it.copy(selected = selected) else it })
        }
    }

    fun renameColumn(key: String, name: String) {
        _state.update { state ->
            state.copy(columns = state.columns.map { if (it.key == key) it.copy(outputName = name) else it })
        }
    }

    fun setTrimWhitespace(value: Boolean) = _state.update { it.copy(trimWhitespace = value) }
    fun setRemoveDuplicates(value: Boolean) = _state.update { it.copy(removeDuplicates = value) }

    fun refreshPreview() {
        val snapshot = _state.value
        if (snapshot.sessionId == null) return
        if (snapshot.columns.none { it.selected }) {
            error("Select at least one output column.")
            return
        }
        viewModelScope.launch {
            _state.update { it.copy(phase = WorkPhase.PREVIEWING, statusMessage = "Building preview…", errorMessage = null) }
            runCatching {
                withContext(Dispatchers.IO) { api.preview(snapshot.baseUrl, buildPayload(snapshot).put("limit", 25)) }
            }.onSuccess { preview ->
                _state.update {
                    it.copy(phase = WorkPhase.IDLE, preview = preview, rowCount = preview.rowCount, statusMessage = "Preview ready.")
                }
            }.onFailure { throwable ->
                _state.update { it.copy(phase = WorkPhase.IDLE, errorMessage = friendlyError(throwable)) }
            }
        }
    }

    fun export(format: String) {
        val snapshot = _state.value
        if (snapshot.sessionId == null) {
            error("Upload files first.")
            return
        }
        if (snapshot.columns.none { it.selected }) {
            error("Select at least one output column.")
            return
        }
        viewModelScope.launch {
            _state.update { it.copy(phase = WorkPhase.EXPORTING, statusMessage = "Creating ${format.uppercase()}…", errorMessage = null) }
            runCatching {
                withContext(Dispatchers.IO) { api.export(snapshot.baseUrl, buildPayload(snapshot).put("format", format), format) }
            }.onSuccess { result ->
                _state.update { it.copy(phase = WorkPhase.IDLE, pendingExport = result, statusMessage = "Choose where to save the export.") }
            }.onFailure { throwable ->
                _state.update { it.copy(phase = WorkPhase.IDLE, errorMessage = friendlyError(throwable)) }
            }
        }
    }

    fun savePendingExport(context: Context, destination: Uri?) {
        val result = _state.value.pendingExport ?: return
        if (destination == null) {
            _state.update { it.copy(pendingExport = null, statusMessage = "Export cancelled.") }
            return
        }
        viewModelScope.launch {
            runCatching {
                withContext(Dispatchers.IO) {
                    context.contentResolver.openOutputStream(destination, "w")?.use { it.write(result.bytes) }
                        ?: throw IllegalStateException("Could not open the selected destination.")
                }
            }.onSuccess {
                _state.update { it.copy(pendingExport = null, statusMessage = "Saved ${result.filename}.", errorMessage = null) }
            }.onFailure { throwable ->
                _state.update { it.copy(pendingExport = null, errorMessage = friendlyError(throwable)) }
            }
        }
    }

    fun dismissError() = _state.update { it.copy(errorMessage = null) }

    private fun buildPayload(state: FormaFlowUiState): JSONObject {
        val columns = JSONArray()
        state.columns.filter { it.selected }.forEach { column ->
            columns.put(JSONObject().put("key", column.key).put("name", column.outputName.ifBlank { column.label }))
        }
        return JSONObject()
            .put("session_id", state.sessionId)
            .put("columns", columns)
            .put("fixed_columns", JSONArray())
            .put("derived_columns", JSONArray())
            .put(
                "options",
                JSONObject()
                    .put("trim_whitespace", state.trimWhitespace)
                    .put("remove_duplicates", state.removeDuplicates)
                    .put("skip_blank_key", JSONObject.NULL)
                    .put("sort_key", JSONObject.NULL)
            )
    }

    private fun queryDocument(context: Context, uri: Uri): Pair<String, Long> {
        var name = uri.lastPathSegment ?: "document"
        var size = -1L
        val cursor: Cursor? = context.contentResolver.query(uri, arrayOf(OpenableColumns.DISPLAY_NAME, OpenableColumns.SIZE), null, null, null)
        cursor?.use {
            if (it.moveToFirst()) {
                val nameIndex = it.getColumnIndex(OpenableColumns.DISPLAY_NAME)
                val sizeIndex = it.getColumnIndex(OpenableColumns.SIZE)
                if (nameIndex >= 0) name = it.getString(nameIndex) ?: name
                if (sizeIndex >= 0 && !it.isNull(sizeIndex)) size = it.getLong(sizeIndex)
            }
        }
        return name to size
    }

    private fun error(message: String) = _state.update { it.copy(errorMessage = message) }

    private fun friendlyError(throwable: Throwable): String = when (throwable) {
        is ApiException -> if (throwable.statusCode == 404) {
            "The temporary server session expired. Upload the files again."
        } else throwable.message ?: "Server error."
        is java.net.SocketTimeoutException -> "The server took too long. A free Render service may still be waking up; try again once."
        is java.net.UnknownHostException -> "Could not reach the server. Check the URL and internet connection."
        else -> throwable.message ?: "Unexpected error."
    }
}
