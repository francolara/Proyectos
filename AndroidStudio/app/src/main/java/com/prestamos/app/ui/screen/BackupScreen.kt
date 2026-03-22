package com.prestamos.app.ui.screen

import android.content.Intent
import android.net.Uri
import androidx.activity.compose.rememberLauncherForActivityResult
import androidx.activity.result.contract.ActivityResultContracts
import androidx.compose.foundation.background
import androidx.compose.foundation.layout.Arrangement
import androidx.compose.foundation.layout.Box
import androidx.compose.foundation.layout.Column
import androidx.compose.foundation.layout.fillMaxSize
import androidx.compose.foundation.layout.fillMaxWidth
import androidx.compose.foundation.layout.height
import androidx.compose.foundation.layout.padding
import androidx.compose.foundation.layout.Row
import androidx.compose.foundation.layout.size
import androidx.compose.foundation.layout.Spacer
import androidx.compose.foundation.rememberScrollState
import androidx.compose.foundation.shape.RoundedCornerShape
import androidx.compose.foundation.verticalScroll
import androidx.compose.material3.AlertDialog
import androidx.compose.material3.Button
import androidx.compose.material3.Card
import androidx.compose.material3.CardDefaults
import androidx.compose.material3.HorizontalDivider
import androidx.compose.material3.OutlinedButton
import androidx.compose.material3.MaterialTheme
import androidx.compose.material3.SnackbarHost
import androidx.compose.material3.SnackbarHostState
import androidx.compose.material3.Text
import androidx.compose.material3.TextButton
import androidx.compose.runtime.Composable
import androidx.compose.runtime.LaunchedEffect
import androidx.compose.runtime.getValue
import androidx.compose.runtime.mutableStateOf
import androidx.compose.runtime.remember
import androidx.compose.runtime.setValue
import androidx.compose.ui.Alignment
import androidx.compose.ui.Modifier
import androidx.compose.ui.graphics.Color
import androidx.compose.ui.platform.LocalContext
import androidx.compose.ui.unit.dp
import androidx.lifecycle.compose.collectAsStateWithLifecycle
import androidx.lifecycle.viewmodel.compose.viewModel
import com.prestamos.app.data.backup.BackupManager
import com.prestamos.app.ui.viewmodel.BackupViewModel
import java.text.SimpleDateFormat
import java.util.Date
import java.util.Locale

@Composable
fun BackupScreen(viewModel: BackupViewModel = viewModel()) {
    val context = LocalContext.current
    val uiState by viewModel.uiState.collectAsStateWithLifecycle()
    val mensaje by viewModel.mensaje.collectAsStateWithLifecycle()
    val snackbarHostState = remember { SnackbarHostState() }
    val scrollState = rememberScrollState()
    var pendingImportUri by remember { mutableStateOf<Uri?>(null) }
    var selectedDestination by remember { mutableStateOf(BackupDestination.LOCAL) }
    val lastBackupTimestamp = uiState.lastBackupTimestamp
    val isLocalSelected = selectedDestination == BackupDestination.LOCAL

    val folderBackupLauncher = rememberLauncherForActivityResult(
        contract = ActivityResultContracts.OpenDocumentTree()
    ) { uri ->
        if (uri != null) {
            viewModel.generarRespaldoEnCarpeta(uri = uri, persistPermission = true)
        } else {
            viewModel.limpiarMensaje()
        }
    }

    val restoreBackupLauncher = rememberLauncherForActivityResult(
        contract = ActivityResultContracts.OpenDocument()
    ) { uri ->
        if (uri != null) {
            viewModel.persistReadPermission(uri)
            pendingImportUri = uri
        }
    }

    LaunchedEffect(Unit) {
        viewModel.refreshLicenseStatus()
    }

    LaunchedEffect(mensaje) {
        val text = mensaje ?: return@LaunchedEffect
        snackbarHostState.showSnackbar(text)
        viewModel.limpiarMensaje()
    }

    Box(modifier = Modifier.fillMaxSize()) {
        Column(
            modifier = Modifier
                .fillMaxSize()
                .verticalScroll(scrollState)
                .padding(16.dp),
            verticalArrangement = Arrangement.spacedBy(12.dp)
        ) {
            Text("Respaldo", style = MaterialTheme.typography.headlineSmall)
            Text(
                "Gestiona copias de seguridad y restauracion de datos",
                style = MaterialTheme.typography.bodyMedium,
                color = MaterialTheme.colorScheme.onSurfaceVariant
            )

            Card(
                shape = RoundedCornerShape(14.dp),
                colors = CardDefaults.cardColors(containerColor = MaterialTheme.colorScheme.secondaryContainer.copy(alpha = 0.45f)),
                modifier = Modifier.fillMaxWidth()
            ) {
                Column(
                    modifier = Modifier
                        .fillMaxWidth()
                        .padding(12.dp),
                    verticalArrangement = Arrangement.spacedBy(4.dp)
                ) {
                    Text("Estado actual", style = MaterialTheme.typography.titleSmall)
                    Text("Archivo: ${BackupManager.BACKUP_FILE_NAME}", style = MaterialTheme.typography.bodySmall)
                    Text(
                        "Ultimo respaldo: ${uiState.lastBackupTimestamp.toDisplayDateTime()}",
                        style = MaterialTheme.typography.bodySmall
                    )
                    Row(verticalAlignment = Alignment.CenterVertically, horizontalArrangement = Arrangement.spacedBy(6.dp)) {
                        Box(
                            modifier = Modifier
                                .size(9.dp)
                                .background(
                                    color = if (uiState.hasSavedLocation) Color(0xFF2E7D32) else Color(0xFFF57F17),
                                    shape = RoundedCornerShape(50)
                                )
                        )
                        Text(
                            text = if (uiState.hasSavedLocation) "Ubicacion configurada" else "Sin ubicacion configurada",
                            style = MaterialTheme.typography.labelMedium
                        )
                    }
                }
            }

            Card(
                shape = RoundedCornerShape(14.dp),
                colors = CardDefaults.cardColors(containerColor = MaterialTheme.colorScheme.surfaceVariant.copy(alpha = 0.35f)),
                modifier = Modifier.fillMaxWidth()
            ) {
                Column(
                    modifier = Modifier
                        .fillMaxWidth()
                        .padding(12.dp),
                    verticalArrangement = Arrangement.spacedBy(8.dp)
                ) {
                    Text("Destino de respaldo", style = MaterialTheme.typography.titleSmall)
                    Row(
                        modifier = Modifier.fillMaxWidth(),
                        horizontalArrangement = Arrangement.spacedBy(8.dp)
                    ) {
                        Button(
                            onClick = { selectedDestination = BackupDestination.LOCAL },
                            modifier = Modifier.weight(1f)
                        ) {
                            Text("Local")
                        }
                        OutlinedButton(
                            onClick = { selectedDestination = BackupDestination.DRIVE },
                            modifier = Modifier.weight(1f)
                        ) {
                            Text("Drive")
                        }
                    }
                    Text(
                        text = if (isLocalSelected) {
                            "Usando respaldo local del dispositivo."
                        } else {
                            "Google Drive estara disponible en una proxima version."
                        },
                        style = MaterialTheme.typography.bodySmall,
                        color = MaterialTheme.colorScheme.onSurfaceVariant
                    )
                }
            }

            if (!uiState.licenseActive) {
                Card(
                    shape = RoundedCornerShape(12.dp),
                    colors = CardDefaults.cardColors(containerColor = Color(0xFFFFE8E8)),
                    modifier = Modifier.fillMaxWidth()
                ) {
                    Text(
                        "Licencia activa requerida para respaldo y restauracion",
                        color = Color(0xFF9D2B2B),
                        style = MaterialTheme.typography.bodyMedium,
                        modifier = Modifier.padding(12.dp)
                    )
                }
            }

            Card(
                shape = RoundedCornerShape(14.dp),
                colors = CardDefaults.cardColors(containerColor = MaterialTheme.colorScheme.secondaryContainer.copy(alpha = 0.35f)),
                modifier = Modifier.fillMaxWidth()
            ) {
                Column(
                    modifier = Modifier
                        .fillMaxWidth()
                        .padding(12.dp),
                    verticalArrangement = Arrangement.spacedBy(8.dp)
                ) {
                    Text("Acciones", style = MaterialTheme.typography.titleSmall)
                    OutlinedButton(
                        modifier = Modifier.fillMaxWidth(),
                        enabled = uiState.licenseActive && isLocalSelected,
                        onClick = { folderBackupLauncher.launch(null) }
                    ) { Text("Configurar ubicacion") }

                    Button(
                        modifier = Modifier.fillMaxWidth(),
                        enabled = uiState.licenseActive && isLocalSelected,
                        onClick = {
                            if (uiState.hasSavedLocation) {
                                viewModel.generarRespaldoEnUbicacionGuardada(
                                    onLocationMissing = { folderBackupLauncher.launch(null) }
                                )
                            } else {
                                folderBackupLauncher.launch(null)
                            }
                        }
                    ) { Text("Crear respaldo ahora") }

                    OutlinedButton(
                        modifier = Modifier.fillMaxWidth(),
                        enabled = uiState.licenseActive && isLocalSelected,
                        onClick = { restoreBackupLauncher.launch(arrayOf("application/json", "text/plain", "*/*")) }
                    ) { Text("Restaurar respaldo") }

                    OutlinedButton(
                        modifier = Modifier.fillMaxWidth(),
                        enabled = uiState.licenseActive && uiState.hasSavedLocation && isLocalSelected,
                        onClick = {
                            viewModel.getSavedBackupUri { uri ->
                                if (uri == null) {
                                    viewModel.reportarError("No hay respaldo para compartir")
                                    return@getSavedBackupUri
                                }
                                val sendIntent = Intent(Intent.ACTION_SEND).apply {
                                    type = "application/json"
                                    putExtra(Intent.EXTRA_STREAM, uri)
                                    putExtra(Intent.EXTRA_SUBJECT, "Respaldo de prestamos")
                                    addFlags(Intent.FLAG_GRANT_READ_URI_PERMISSION)
                                }
                                runCatching {
                                    context.startActivity(Intent.createChooser(sendIntent, "Compartir respaldo"))
                                }.onFailure {
                                    viewModel.reportarError("Error al compartir respaldo")
                                }
                            }
                        }
                    ) { Text("Compartir archivo de respaldo") }

                    OutlinedButton(
                        modifier = Modifier.fillMaxWidth(),
                        enabled = uiState.licenseActive && uiState.hasSavedLocation && isLocalSelected,
                        onClick = {
                            viewModel.getSavedBackupUri { uri ->
                                if (uri == null) {
                                    viewModel.reportarError("No hay ubicacion configurada")
                                    return@getSavedBackupUri
                                }
                                val openIntent = Intent(Intent.ACTION_VIEW).apply {
                                    setData(uri)
                                    addFlags(Intent.FLAG_GRANT_READ_URI_PERMISSION)
                                    addFlags(Intent.FLAG_ACTIVITY_NEW_TASK)
                                }
                                runCatching { context.startActivity(openIntent) }
                                    .onFailure { viewModel.reportarError("No se pudo abrir la ubicacion") }
                            }
                        }
                    ) { Text("Abrir ubicacion de respaldo") }
                }
            }

            Card(
                shape = RoundedCornerShape(14.dp),
                colors = CardDefaults.cardColors(containerColor = MaterialTheme.colorScheme.surfaceVariant.copy(alpha = 0.35f)),
                modifier = Modifier.fillMaxWidth()
            ) {
                Column(
                    modifier = Modifier
                        .fillMaxWidth()
                        .padding(12.dp),
                    verticalArrangement = Arrangement.spacedBy(8.dp)
                ) {
                    Text("Historial reciente", style = MaterialTheme.typography.titleSmall)
                    if (lastBackupTimestamp == null || lastBackupTimestamp <= 0L) {
                        Text("Aun no hay respaldos registrados.", style = MaterialTheme.typography.bodySmall)
                    } else {
                        Text("Respaldo principal", style = MaterialTheme.typography.labelLarge)
                        Text("Fecha: ${lastBackupTimestamp.toDisplayDateTime()}", style = MaterialTheme.typography.bodySmall)
                        Text("Archivo: ${BackupManager.BACKUP_FILE_NAME}", style = MaterialTheme.typography.bodySmall)
                        HorizontalDivider()
                        Text(
                            "Tip: crea un respaldo antes de restaurar para mantener un punto de retorno.",
                            style = MaterialTheme.typography.labelSmall,
                            color = MaterialTheme.colorScheme.onSurfaceVariant
                        )
                    }
                }
            }

            Spacer(modifier = Modifier.height(56.dp))
        }

        SnackbarHost(
            hostState = snackbarHostState,
            modifier = Modifier
                .align(Alignment.BottomCenter)
                .padding(12.dp)
        )
    }

    pendingImportUri?.let { uri ->
        AlertDialog(
            onDismissRequest = { pendingImportUri = null },
            title = { Text("Confirmar restauracion") },
            text = { Text("Esta accion reemplazara todos los datos actuales. Desea continuar?") },
            confirmButton = {
                TextButton(onClick = {
                    pendingImportUri = null
                    viewModel.restaurarRespaldo(uri) {
                        BackupManager.restartApplication(context)
                    }
                }) {
                    Text("Restaurar")
                }
            },
            dismissButton = {
                TextButton(onClick = { pendingImportUri = null }) {
                    Text("Cancelar")
                }
            }
        )
    }
}

private enum class BackupDestination {
    LOCAL,
    DRIVE
}

private fun Long?.toDisplayDateTime(): String {
    if (this == null || this <= 0L) return "Nunca"
    val formatter = SimpleDateFormat("dd/MM/yyyy HH:mm", Locale.getDefault())
    return formatter.format(Date(this))
}
