package uz.formaflow.mobile.data

import android.content.Context
import org.json.JSONArray
import org.json.JSONObject
import java.io.BufferedInputStream
import java.io.BufferedOutputStream
import java.net.HttpURLConnection
import java.net.URL
import java.nio.charset.StandardCharsets
import java.util.UUID

class FormaFlowApi {
    private val connectTimeoutMs = 120_000
    private val readTimeoutMs = 360_000

    fun health(baseUrl: String): HealthResult {
        val connection = open(baseUrl, "/health", "GET")
        return connection.useConnection { code, body, _ ->
            ensureSuccess(code, body)
            val json = JSONObject(body.toString(StandardCharsets.UTF_8))
            HealthResult(json.optString("status", "unknown"), json.optString("version", "?"))
        }
    }

    fun upload(
        context: Context,
        baseUrl: String,
        files: List<PickedDocument>,
        onProgress: (Float) -> Unit,
    ): UploadResponse {
        val boundary = "FormaFlow-${UUID.randomUUID()}"
        val connection = open(baseUrl, "/api/upload", "POST").apply {
            doOutput = true
            setRequestProperty("Accept", "application/json")
            setRequestProperty("Content-Type", "multipart/form-data; boundary=$boundary")
            setChunkedStreamingMode(64 * 1024)
        }

        val total = files.sumOf { it.sizeBytes.coerceAtLeast(0L) }.coerceAtLeast(1L)
        var sent = 0L
        val resolver = context.contentResolver

        BufferedOutputStream(connection.outputStream).use { output ->
            files.forEach { file ->
                output.writeUtf8("--$boundary\r\n")
                output.writeUtf8(
                    "Content-Disposition: form-data; name=\"files\"; filename=\"${safeHeader(file.name)}\"\r\n"
                )
                output.writeUtf8("Content-Type: ${file.mimeType.ifBlank { "application/octet-stream" }}\r\n\r\n")

                resolver.openInputStream(file.uri)?.use { raw ->
                    BufferedInputStream(raw).use { input ->
                        val buffer = ByteArray(64 * 1024)
                        while (true) {
                            val count = input.read(buffer)
                            if (count < 0) break
                            output.write(buffer, 0, count)
                            sent += count
                            onProgress((sent.toDouble() / total.toDouble()).toFloat().coerceIn(0f, 1f))
                        }
                    }
                } ?: throw ApiException("Could not open ${file.name}.")
                output.writeUtf8("\r\n")
            }
            output.writeUtf8("--$boundary--\r\n")
            output.flush()
        }

        return connection.useConnection { code, body, _ ->
            ensureSuccess(code, body)
            parseUpload(JSONObject(body.toString(StandardCharsets.UTF_8)))
        }
    }

    fun selectDataset(baseUrl: String, sessionId: String, datasetId: String): SelectedDatasetResponse {
        val payload = JSONObject()
            .put("session_id", sessionId)
            .put("dataset_id", datasetId)
        val json = postJson(baseUrl, "/api/select-dataset", payload)
        return SelectedDatasetResponse(
            activeDatasetId = json.getString("active_dataset_id"),
            columns = parseColumns(json.optJSONArray("columns") ?: JSONArray()),
            rowCount = json.optInt("row_count"),
            datasetType = json.optString("dataset_type"),
        )
    }

    fun preview(baseUrl: String, payload: JSONObject): PreviewResponse {
        val json = postJson(baseUrl, "/api/preview", payload)
        val headersArray = json.optJSONArray("headers") ?: JSONArray()
        val headers = List(headersArray.length()) { headersArray.optString(it) }
        val rowsArray = json.optJSONArray("rows") ?: JSONArray()
        val rows = buildList {
            for (index in 0 until rowsArray.length()) {
                val rowJson = rowsArray.optJSONObject(index) ?: continue
                val row = linkedMapOf<String, String>()
                headers.forEach { header ->
                    val value = rowJson.opt(header)
                    row[header] = when (value) {
                        null, JSONObject.NULL -> ""
                        else -> value.toString()
                    }
                }
                add(row)
            }
        }
        return PreviewResponse(
            headers = headers,
            rows = rows,
            rowCount = json.optInt("row_count"),
            previewCount = json.optInt("preview_count"),
        )
    }

    fun export(baseUrl: String, payload: JSONObject, fallbackFormat: String): ExportResult {
        val connection = open(baseUrl, "/api/export", "POST").apply {
            doOutput = true
            setRequestProperty("Accept", "*/*")
            setRequestProperty("Content-Type", "application/json; charset=utf-8")
        }
        connection.outputStream.use { it.write(payload.toString().toByteArray(StandardCharsets.UTF_8)) }
        return connection.useConnection { code, body, headers ->
            ensureSuccess(code, body)
            val disposition = headers.firstHeader("Content-Disposition").orEmpty()
            val filename = Regex("filename=\\\"?([^\\\";]+)").find(disposition)?.groupValues?.getOrNull(1)
                ?: "FormaFlow_Output.$fallbackFormat"
            val mime = headers["Content-Type"]?.firstOrNull() ?: mimeFor(fallbackFormat)
            ExportResult(body, filename, mime)
        }
    }

    private fun postJson(baseUrl: String, path: String, payload: JSONObject): JSONObject {
        val connection = open(baseUrl, path, "POST").apply {
            doOutput = true
            setRequestProperty("Accept", "application/json")
            setRequestProperty("Content-Type", "application/json; charset=utf-8")
        }
        connection.outputStream.use { it.write(payload.toString().toByteArray(StandardCharsets.UTF_8)) }
        return connection.useConnection { code, body, _ ->
            ensureSuccess(code, body)
            JSONObject(body.toString(StandardCharsets.UTF_8))
        }
    }

    private fun parseUpload(json: JSONObject): UploadResponse {
        val filesArray = json.optJSONArray("files") ?: JSONArray()
        val files = buildList {
            for (index in 0 until filesArray.length()) {
                val item = filesArray.optJSONObject(index) ?: continue
                add(
                    FileSummary(
                        name = item.optString("name"),
                        ok = item.optBoolean("ok"),
                        error = item.optNullableString("error"),
                        detectedType = item.optString("detected_type", "Document"),
                        documentType = item.optString("document_type", "unknown_document"),
                        rows = item.optInt("rows"),
                        columns = item.optInt("columns"),
                        warnings = item.optJSONArray("warnings").toStringList(),
                    )
                )
            }
        }
        val groupsArray = json.optJSONArray("dataset_groups") ?: JSONArray()
        val groups = buildList {
            for (index in 0 until groupsArray.length()) {
                val item = groupsArray.optJSONObject(index) ?: continue
                add(
                    DatasetGroup(
                        id = item.optString("id"),
                        name = item.optString("name"),
                        datasetType = item.optString("dataset_type"),
                        rowCount = item.optInt("row_count"),
                        sourceCount = item.optInt("source_count"),
                        tableCount = item.optInt("table_count"),
                        sources = item.optJSONArray("sources").toStringList(),
                    )
                )
            }
        }
        return UploadResponse(
            sessionId = json.getString("session_id"),
            files = files,
            datasetGroups = groups,
            activeDatasetId = json.optNullableString("active_dataset_id"),
            columns = parseColumns(json.optJSONArray("columns") ?: JSONArray()),
            rowCount = json.optInt("row_count"),
            message = json.optNullableString("message"),
        )
    }

    private fun parseColumns(array: JSONArray): List<ColumnInfo> = buildList {
        for (index in 0 until array.length()) {
            val item = array.optJSONObject(index) ?: continue
            val label = item.optString("label")
            val original = item.optString("original").ifBlank {
                item.optJSONArray("originals")?.optString(0).orEmpty()
            }
            add(
                ColumnInfo(
                    key = item.optString("key"),
                    label = label,
                    original = original,
                    confidence = item.optDouble("confidence", 0.0),
                    method = item.optString("method", "custom"),
                )
            )
        }
    }

    private fun open(baseUrl: String, path: String, method: String): HttpURLConnection {
        val normalized = normalizeBaseUrl(baseUrl)
        return (URL(normalized + path.removePrefix("/")).openConnection() as HttpURLConnection).apply {
            requestMethod = method
            connectTimeout = connectTimeoutMs
            readTimeout = readTimeoutMs
            useCaches = false
            instanceFollowRedirects = true
            setRequestProperty("User-Agent", "FormaFlow-Android/0.1")
        }
    }

    private inline fun <T> HttpURLConnection.useConnection(
        block: (code: Int, body: ByteArray, headers: Map<String, List<String>>) -> T,
    ): T {
        try {
            val code = responseCode
            val stream = if (code in 200..299) inputStream else errorStream
            val bytes = stream?.use { it.readBytes() } ?: ByteArray(0)
            val headers = headerFields.entries
                .filter { it.key != null }
                .associate { it.key!! to it.value }
            return block(code, bytes, headers)
        } finally {
            disconnect()
        }
    }

    private fun ensureSuccess(code: Int, body: ByteArray) {
        if (code in 200..299) return
        val text = body.toString(StandardCharsets.UTF_8)
        val message = runCatching {
            val detail = JSONObject(text).opt("detail")
            when (detail) {
                is String -> detail
                null, JSONObject.NULL -> text
                else -> detail.toString()
            }
        }.getOrDefault(text).ifBlank { "Server returned HTTP $code." }
        throw ApiException(message, code)
    }

    companion object {
        fun normalizeBaseUrl(value: String): String {
            val trimmed = value.trim()
            val withScheme = if (trimmed.startsWith("http://") || trimmed.startsWith("https://")) {
                trimmed
            } else {
                "https://$trimmed"
            }
            require(withScheme.startsWith("https://")) { "Use an HTTPS server URL." }
            return withScheme.trimEnd('/') + "/"
        }

        private fun safeHeader(value: String): String = value.replace("\"", "'").replace("\r", " ").replace("\n", " ")
        private fun mimeFor(format: String): String = when (format.lowercase()) {
            "xlsx" -> "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            "csv" -> "text/csv"
            "docx" -> "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
            else -> "application/octet-stream"
        }
    }
}

private fun BufferedOutputStream.writeUtf8(value: String) {
    write(value.toByteArray(StandardCharsets.UTF_8))
}

private fun JSONObject.optNullableString(key: String): String? {
    if (!has(key) || isNull(key)) return null
    return optString(key).takeIf { it.isNotBlank() && it != "null" }
}

private fun JSONArray?.toStringList(): List<String> {
    if (this == null) return emptyList()
    return List(length()) { optString(it) }.filter { it.isNotBlank() }
}

private fun Map<String, List<String>>.firstHeader(name: String): String? =
    entries.firstOrNull { it.key.equals(name, ignoreCase = true) }?.value?.firstOrNull()
