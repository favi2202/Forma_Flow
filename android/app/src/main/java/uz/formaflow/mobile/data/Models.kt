package uz.formaflow.mobile.data

import android.net.Uri

data class PickedDocument(
    val uri: Uri,
    val name: String,
    val sizeBytes: Long,
    val mimeType: String,
)

data class HealthResult(val status: String, val version: String)

data class ColumnInfo(
    val key: String,
    val label: String,
    val original: String,
    val confidence: Double,
    val method: String,
    val outputName: String = label.ifBlank { original.ifBlank { key } },
    val selected: Boolean = true,
)

data class FileSummary(
    val name: String,
    val ok: Boolean,
    val error: String?,
    val detectedType: String,
    val documentType: String,
    val rows: Int,
    val columns: Int,
    val warnings: List<String>,
)

data class DatasetGroup(
    val id: String,
    val name: String,
    val datasetType: String,
    val rowCount: Int,
    val sourceCount: Int,
    val tableCount: Int,
    val sources: List<String>,
)

data class UploadResponse(
    val sessionId: String,
    val files: List<FileSummary>,
    val datasetGroups: List<DatasetGroup>,
    val activeDatasetId: String?,
    val columns: List<ColumnInfo>,
    val rowCount: Int,
    val message: String?,
)

data class SelectedDatasetResponse(
    val activeDatasetId: String,
    val columns: List<ColumnInfo>,
    val rowCount: Int,
    val datasetType: String,
)

data class PreviewResponse(
    val headers: List<String>,
    val rows: List<Map<String, String>>,
    val rowCount: Int,
    val previewCount: Int,
)

data class ExportResult(
    val bytes: ByteArray,
    val filename: String,
    val mimeType: String,
)

class ApiException(message: String, val statusCode: Int? = null) : Exception(message)
