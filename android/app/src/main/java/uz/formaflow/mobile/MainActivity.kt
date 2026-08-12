package uz.formaflow.mobile

import android.content.Intent
import android.os.Bundle
import androidx.activity.ComponentActivity
import androidx.activity.compose.rememberLauncherForActivityResult
import androidx.activity.compose.setContent
import androidx.activity.result.contract.ActivityResultContracts
import androidx.activity.viewModels
import androidx.compose.foundation.background
import androidx.compose.foundation.horizontalScroll
import androidx.compose.foundation.layout.Arrangement
import androidx.compose.foundation.layout.Box
import androidx.compose.foundation.layout.Column
import androidx.compose.foundation.layout.Row
import androidx.compose.foundation.layout.Spacer
import androidx.compose.foundation.layout.fillMaxSize
import androidx.compose.foundation.layout.fillMaxWidth
import androidx.compose.foundation.layout.height
import androidx.compose.foundation.layout.padding
import androidx.compose.foundation.layout.width
import androidx.compose.foundation.lazy.LazyColumn
import androidx.compose.foundation.lazy.items
import androidx.compose.foundation.rememberScrollState
import androidx.compose.material3.AlertDialog
import androidx.compose.material3.Button
import androidx.compose.material3.Card
import androidx.compose.material3.CardDefaults
import androidx.compose.material3.Checkbox
import androidx.compose.material3.DropdownMenu
import androidx.compose.material3.DropdownMenuItem
import androidx.compose.material3.ExperimentalMaterial3Api
import androidx.compose.material3.HorizontalDivider
import androidx.compose.material3.LinearProgressIndicator
import androidx.compose.material3.MaterialTheme
import androidx.compose.material3.OutlinedButton
import androidx.compose.material3.OutlinedTextField
import androidx.compose.material3.Scaffold
import androidx.compose.material3.Switch
import androidx.compose.material3.Text
import androidx.compose.material3.TextButton
import androidx.compose.material3.TopAppBar
import androidx.compose.runtime.Composable
import androidx.compose.runtime.LaunchedEffect
import androidx.compose.runtime.collectAsState
import androidx.compose.runtime.getValue
import androidx.compose.runtime.mutableStateOf
import androidx.compose.runtime.remember
import androidx.compose.runtime.setValue
import androidx.compose.ui.Alignment
import androidx.compose.ui.Modifier
import androidx.compose.ui.graphics.Color
import androidx.compose.ui.text.font.FontWeight
import androidx.compose.ui.text.style.TextOverflow
import androidx.compose.ui.unit.dp
import uz.formaflow.mobile.data.ColumnInfo
import uz.formaflow.mobile.data.DatasetGroup
import uz.formaflow.mobile.data.FileSummary
import uz.formaflow.mobile.data.PickedDocument
import uz.formaflow.mobile.data.PreviewResponse
import uz.formaflow.mobile.ui.FormaFlowUiState
import uz.formaflow.mobile.ui.FormaFlowViewModel
import uz.formaflow.mobile.ui.ServerState
import uz.formaflow.mobile.ui.WorkPhase
import uz.formaflow.mobile.ui.theme.FormaFlowTheme

class MainActivity : ComponentActivity() {
    private val viewModel: FormaFlowViewModel by viewModels()

    override fun onCreate(savedInstanceState: Bundle?) {
        super.onCreate(savedInstanceState)
        setContent {
            FormaFlowTheme {
                FormaFlowScreen(viewModel)
            }
        }
    }
}

@OptIn(ExperimentalMaterial3Api::class)
@Composable
private fun FormaFlowScreen(viewModel: FormaFlowViewModel) {
    val state by viewModel.state.collectAsState()
    val context = androidx.compose.ui.platform.LocalContext.current

    val documentPicker = rememberLauncherForActivityResult(ActivityResultContracts.OpenMultipleDocuments()) { uris ->
        uris.forEach { uri ->
            runCatching {
                context.contentResolver.takePersistableUriPermission(uri, Intent.FLAG_GRANT_READ_URI_PERMISSION)
            }
        }
        viewModel.addDocuments(context, uris)
    }
    val exportPicker = rememberLauncherForActivityResult(ActivityResultContracts.CreateDocument("application/octet-stream")) { uri ->
        viewModel.savePendingExport(context, uri)
    }
    LaunchedEffect(state.pendingExport?.filename) {
        state.pendingExport?.let { exportPicker.launch(it.filename) }
    }

    Scaffold(
        topBar = {
            TopAppBar(
                title = {
                    Column {
                        Text("FormaFlow", fontWeight = FontWeight.Bold)
                        Text("Native Android · v0.1", style = MaterialTheme.typography.labelSmall)
                    }
                },
                actions = {
                    StatusPill(state.serverState, state.serverVersion)
                    Spacer(Modifier.width(12.dp))
                },
            )
        }
    ) { padding ->
        LazyColumn(
            modifier = Modifier.fillMaxSize().padding(padding),
            contentPadding = androidx.compose.foundation.layout.PaddingValues(16.dp),
            verticalArrangement = Arrangement.spacedBy(14.dp),
        ) {
            item { ServerCard(state, viewModel) }
            item { PrivacyCard() }
            item {
                FilePickerCard(
                    state = state,
                    onPick = {
                        documentPicker.launch(arrayOf("*/*"))
                    },
                    onClear = viewModel::clearAll,
                    onUpload = { viewModel.uploadAndAnalyze(context) },
                )
            }
            if (state.fileSummaries.isNotEmpty()) item { ResultsCard(state.fileSummaries) }
            if (state.datasetGroups.isNotEmpty()) item {
                DatasetCard(state.datasetGroups, state.activeDatasetId, viewModel::selectDataset)
            }
            if (state.columns.isNotEmpty()) {
                item {
                    Text("Output columns", style = MaterialTheme.typography.titleLarge, fontWeight = FontWeight.Bold)
                    Text("Choose fields and rename the exported headings.", color = MaterialTheme.colorScheme.onSurfaceVariant)
                }
                items(state.columns, key = { it.key }) { column ->
                    ColumnCard(column, viewModel::toggleColumn, viewModel::renameColumn)
                }
                item { OptionsCard(state, viewModel) }
                item { PreviewCard(state.preview, state.rowCount, state.phase, viewModel::refreshPreview) }
                item { ExportCard(state.phase, viewModel::export) }
                item {
                    Text(
                        "Temporary server sessions expire after 30 minutes of inactivity. Clear removes the active session immediately.",
                        style = MaterialTheme.typography.bodySmall,
                        color = MaterialTheme.colorScheme.onSurfaceVariant,
                    )
                }
            }
            item { Spacer(Modifier.height(40.dp)) }
        }
    }

    state.errorMessage?.let { message ->
        AlertDialog(
            onDismissRequest = viewModel::dismissError,
            confirmButton = { TextButton(onClick = viewModel::dismissError) { Text("OK") } },
            title = { Text("FormaFlow") },
            text = { Text(message) },
        )
    }
}

@Composable
private fun StatusPill(status: ServerState, version: String) {
    val (label, color) = when (status) {
        ServerState.READY -> "Ready${if (version.isNotBlank()) " · $version" else ""}" to Color(0xFF0B7A4B)
        ServerState.CHECKING -> "Checking" to Color(0xFF8A6100)
        ServerState.WAKING -> "Waking" to Color(0xFF8A6100)
        ServerState.OFFLINE -> "Offline" to MaterialTheme.colorScheme.error
        ServerState.UNKNOWN -> "Unknown" to MaterialTheme.colorScheme.onSurfaceVariant
    }
    Box(
        modifier = Modifier.background(color.copy(alpha = 0.12f), MaterialTheme.shapes.large).padding(horizontal = 10.dp, vertical = 6.dp)
    ) {
        Text(label, color = color, style = MaterialTheme.typography.labelMedium)
    }
}

@Composable
private fun ServerCard(state: FormaFlowUiState, viewModel: FormaFlowViewModel) {
    Card {
        Column(Modifier.padding(16.dp), verticalArrangement = Arrangement.spacedBy(10.dp)) {
            Text("Server connection", style = MaterialTheme.typography.titleMedium, fontWeight = FontWeight.Bold)
            OutlinedTextField(
                value = state.baseUrl,
                onValueChange = viewModel::setBaseUrl,
                label = { Text("FormaFlow HTTPS URL") },
                modifier = Modifier.fillMaxWidth(),
                singleLine = true,
            )
            Row(
                modifier = Modifier.horizontalScroll(rememberScrollState()),
                horizontalArrangement = Arrangement.spacedBy(8.dp),
            ) {
                Button(onClick = viewModel::saveBaseUrl, enabled = state.phase == WorkPhase.IDLE) { Text("Save and test") }
                OutlinedButton(onClick = viewModel::checkServer, enabled = state.phase == WorkPhase.IDLE) { Text("Test again") }
            }
            if (state.statusMessage.isNotBlank()) Text(state.statusMessage, style = MaterialTheme.typography.bodySmall)
        }
    }
}

@Composable
private fun PrivacyCard() {
    Card(colors = CardDefaults.cardColors(containerColor = MaterialTheme.colorScheme.secondaryContainer.copy(alpha = 0.55f))) {
        Column(Modifier.padding(16.dp), verticalArrangement = Arrangement.spacedBy(6.dp)) {
            Text("Privacy notice", style = MaterialTheme.typography.titleMedium, fontWeight = FontWeight.Bold)
            Text(
                "Files selected here are sent to the configured Render server for analysis. During testing, avoid real PINFL/JSHSHIR, phone numbers, addresses, or other sensitive student data.",
                style = MaterialTheme.typography.bodyMedium,
            )
        }
    }
}

@Composable
private fun FilePickerCard(
    state: FormaFlowUiState,
    onPick: () -> Unit,
    onClear: () -> Unit,
    onUpload: () -> Unit,
) {
    val busy = state.phase != WorkPhase.IDLE
    Card {
        Column(Modifier.padding(16.dp), verticalArrangement = Arrangement.spacedBy(12.dp)) {
            Text("Files", style = MaterialTheme.typography.titleMedium, fontWeight = FontWeight.Bold)
            Row(
                modifier = Modifier.horizontalScroll(rememberScrollState()),
                horizontalArrangement = Arrangement.spacedBy(8.dp),
            ) {
                Button(onClick = onPick, enabled = !busy) { Text("Choose files") }
                OutlinedButton(onClick = onClear, enabled = !busy && state.documents.isNotEmpty()) { Text("Clear") }
            }
            if (state.documents.isEmpty()) {
                Text("Excel, CSV, Word, and PDF · up to 25 MB each", color = MaterialTheme.colorScheme.onSurfaceVariant)
            } else {
                state.documents.forEach { DocumentRow(it) }
                Button(onClick = onUpload, enabled = !busy, modifier = Modifier.fillMaxWidth()) {
                    Text("Upload and analyze ${state.documents.size} file(s)")
                }
            }
            if (busy) {
                if (state.phase == WorkPhase.UPLOADING) {
                    LinearProgressIndicator(progress = { state.progress }, modifier = Modifier.fillMaxWidth())
                    Text("Uploading ${(state.progress * 100).toInt()}%")
                } else {
                    LinearProgressIndicator(modifier = Modifier.fillMaxWidth())
                    Text(
                        when (state.phase) {
                            WorkPhase.ANALYZING -> "Analyzing on Render…"
                            WorkPhase.PREVIEWING -> "Building preview…"
                            WorkPhase.EXPORTING -> "Creating export…"
                            else -> "Working…"
                        }
                    )
                }
            }
        }
    }
}

@Composable
private fun DocumentRow(document: PickedDocument) {
    Row(Modifier.fillMaxWidth(), verticalAlignment = Alignment.CenterVertically) {
        Column(Modifier.weight(1f)) {
            Text(document.name, maxLines = 1, overflow = TextOverflow.Ellipsis, fontWeight = FontWeight.SemiBold)
            Text(formatBytes(document.sizeBytes), style = MaterialTheme.typography.bodySmall, color = MaterialTheme.colorScheme.onSurfaceVariant)
        }
    }
}

@Composable
private fun ResultsCard(results: List<FileSummary>) {
    Card {
        Column(Modifier.padding(16.dp), verticalArrangement = Arrangement.spacedBy(10.dp)) {
            Text("Analysis results", style = MaterialTheme.typography.titleMedium, fontWeight = FontWeight.Bold)
            results.forEachIndexed { index, result ->
                if (index > 0) HorizontalDivider()
                Column(verticalArrangement = Arrangement.spacedBy(3.dp)) {
                    Text(result.name, fontWeight = FontWeight.SemiBold)
                    if (result.ok) {
                        Text("${result.detectedType} · ${result.documentType.replace('_', ' ')} · ${result.rows} rows · ${result.columns} columns")
                        result.warnings.take(3).forEach { Text("• $it", style = MaterialTheme.typography.bodySmall) }
                    } else {
                        Text(result.error ?: "Could not parse this file.", color = MaterialTheme.colorScheme.error)
                    }
                }
            }
        }
    }
}

@OptIn(ExperimentalMaterial3Api::class)
@Composable
private fun DatasetCard(groups: List<DatasetGroup>, activeId: String?, onSelect: (String) -> Unit) {
    var expanded by remember { mutableStateOf(false) }
    val active = groups.firstOrNull { it.id == activeId } ?: groups.first()
    Card {
        Column(Modifier.padding(16.dp), verticalArrangement = Arrangement.spacedBy(10.dp)) {
            Text("Detected dataset", style = MaterialTheme.typography.titleMedium, fontWeight = FontWeight.Bold)
            Box {
                OutlinedButton(
                    onClick = { expanded = true },
                    modifier = Modifier.fillMaxWidth(),
                ) {
                    Text("${active.name} · ${active.rowCount} rows", modifier = Modifier.weight(1f), maxLines = 1)
                    Text("▾")
                }
                DropdownMenu(expanded = expanded, onDismissRequest = { expanded = false }) {
                    groups.forEach { group ->
                        DropdownMenuItem(
                            text = { Text("${group.name} · ${group.rowCount} rows · ${group.sourceCount} source(s)") },
                            onClick = {
                                expanded = false
                                onSelect(group.id)
                            },
                        )
                    }
                }
            }
            Text(active.sources.joinToString(" · "), style = MaterialTheme.typography.bodySmall, color = MaterialTheme.colorScheme.onSurfaceVariant)
        }
    }
}

@Composable
private fun ColumnCard(
    column: ColumnInfo,
    onToggle: (String, Boolean) -> Unit,
    onRename: (String, String) -> Unit,
) {
    Card(colors = CardDefaults.cardColors(containerColor = if (column.selected) MaterialTheme.colorScheme.primaryContainer.copy(alpha = 0.35f) else MaterialTheme.colorScheme.surface)) {
        Row(Modifier.fillMaxWidth().padding(12.dp), verticalAlignment = Alignment.CenterVertically) {
            Checkbox(checked = column.selected, onCheckedChange = { onToggle(column.key, it) })
            Column(Modifier.weight(1f), verticalArrangement = Arrangement.spacedBy(4.dp)) {
                OutlinedTextField(
                    value = column.outputName,
                    onValueChange = { onRename(column.key, it) },
                    modifier = Modifier.fillMaxWidth(),
                    label = { Text("Export heading") },
                    singleLine = true,
                    enabled = column.selected,
                )
                Text(
                    "Source: ${column.original.ifBlank { column.key }} · ${(column.confidence * 100).toInt()}% · ${column.method}",
                    style = MaterialTheme.typography.bodySmall,
                    color = MaterialTheme.colorScheme.onSurfaceVariant,
                )
            }
        }
    }
}

@Composable
private fun OptionsCard(state: FormaFlowUiState, viewModel: FormaFlowViewModel) {
    Card {
        Column(Modifier.padding(16.dp), verticalArrangement = Arrangement.spacedBy(10.dp)) {
            Text("Processing options", style = MaterialTheme.typography.titleMedium, fontWeight = FontWeight.Bold)
            SwitchRow("Trim whitespace", state.trimWhitespace, viewModel::setTrimWhitespace)
            SwitchRow("Remove duplicate rows", state.removeDuplicates, viewModel::setRemoveDuplicates)
        }
    }
}

@Composable
private fun SwitchRow(label: String, checked: Boolean, onChange: (Boolean) -> Unit) {
    Row(Modifier.fillMaxWidth(), verticalAlignment = Alignment.CenterVertically) {
        Text(label, modifier = Modifier.weight(1f))
        Switch(checked = checked, onCheckedChange = onChange)
    }
}

@Composable
private fun PreviewCard(preview: PreviewResponse?, rowCount: Int, phase: WorkPhase, onRefresh: () -> Unit) {
    Card {
        Column(Modifier.padding(16.dp), verticalArrangement = Arrangement.spacedBy(10.dp)) {
            Row(Modifier.fillMaxWidth(), verticalAlignment = Alignment.CenterVertically) {
                Column(Modifier.weight(1f)) {
                    Text("Preview", style = MaterialTheme.typography.titleMedium, fontWeight = FontWeight.Bold)
                    Text("$rowCount total rows", style = MaterialTheme.typography.bodySmall)
                }
                OutlinedButton(onClick = onRefresh, enabled = phase == WorkPhase.IDLE) { Text("Refresh") }
            }
            if (preview == null) {
                Text("No preview yet.", color = MaterialTheme.colorScheme.onSurfaceVariant)
            } else if (preview.headers.isEmpty()) {
                Text("No rows available.")
            } else {
                val scroll = rememberScrollState()
                Column(Modifier.horizontalScroll(scroll)) {
                    PreviewRow(preview.headers, header = true)
                    preview.rows.forEach { row -> PreviewRow(preview.headers.map { row[it].orEmpty() }, header = false) }
                }
                Text("Showing ${preview.previewCount} of ${preview.rowCount} rows", style = MaterialTheme.typography.bodySmall)
            }
        }
    }
}

@Composable
private fun PreviewRow(values: List<String>, header: Boolean) {
    Row {
        values.forEach { value ->
            Box(
                modifier = Modifier
                    .width(150.dp)
                    .background(if (header) MaterialTheme.colorScheme.primaryContainer else Color.Transparent)
                    .padding(8.dp)
            ) {
                Text(value, fontWeight = if (header) FontWeight.Bold else FontWeight.Normal, maxLines = 3, overflow = TextOverflow.Ellipsis)
            }
        }
    }
    HorizontalDivider()
}

@Composable
private fun ExportCard(phase: WorkPhase, onExport: (String) -> Unit) {
    Card {
        Column(Modifier.padding(16.dp), verticalArrangement = Arrangement.spacedBy(10.dp)) {
            Text("Export", style = MaterialTheme.typography.titleMedium, fontWeight = FontWeight.Bold)
            Text("Android will ask where to save the generated file.", color = MaterialTheme.colorScheme.onSurfaceVariant)
            Row(
                modifier = Modifier.horizontalScroll(rememberScrollState()),
                horizontalArrangement = Arrangement.spacedBy(8.dp),
            ) {
                Button(onClick = { onExport("xlsx") }, enabled = phase == WorkPhase.IDLE) { Text("Excel") }
                Button(onClick = { onExport("csv") }, enabled = phase == WorkPhase.IDLE) { Text("CSV") }
                Button(onClick = { onExport("docx") }, enabled = phase == WorkPhase.IDLE) { Text("Word") }
            }
        }
    }
}

private fun formatBytes(bytes: Long): String = when {
    bytes < 0 -> "Unknown size"
    bytes < 1024 -> "$bytes B"
    bytes < 1024 * 1024 -> "%.1f KB".format(bytes / 1024.0)
    else -> "%.1f MB".format(bytes / (1024.0 * 1024.0))
}
