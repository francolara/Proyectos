package com.prestamos.app.ui.screen

import android.content.Intent
import android.net.Uri
import android.provider.DocumentsContract
import androidx.activity.compose.rememberLauncherForActivityResult
import androidx.activity.result.ActivityResult
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
import com.prestamos.app.data.backup.BackupStorageDestination
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
    val isLocalSelected = selectedDestination == BackupDestination.LOCAL
    val storageDestination = if (isLocalSelected) BackupStorageDestination.LOCAL else BackupStorageDestination.DRIVE
    val hasSavedLocation = if (isLocalSelected) uiState.hasSavedLocationLocal else uiState.hasSavedLocationDrive

    val folderBackupLauncher = rememberLauncherForActivityResult(
        contract = ActivityResultContracts.StartActivityForResult()
    ) { result: ActivityResult ->
        val uri = result.data?.data
        if (uri != null) {
            viewModel.generarRespaldoEnCarpeta(
                uri = uri,
                persistPermission = true,
                destination = storageDestination
            )
        } else {
            viewModel.limpiarMensaje()
        }
    }

    val driveFileBackupLauncher = rememberLauncherForActivityResult(
        contract = ActivityResultContracts.StartActivityForResult()
    ) { result: ActivityResult ->
        val uri = result.data?.data
        if (uri != null) {
            viewModel.generarRespaldoEnUri(
                uri = uri,
                persistPermission = true,
                destination = BackupStorageDestination.DRIVE
            )
        } else {
            viewModel.limpiarMensaje()
        }
    }

    val driveExistingFileLauncher = rememberLauncherForActivityResult(
        contract = ActivityResultContracts.StartActivityForResult()
    ) { result: ActivityResult ->
        val uri = result.data?.data
        if (uri != null) {
            viewModel.persistReadPermission(uri)
            viewModel.generarRespaldoEnUri(
                uri = uri,
                persistPermission = true,
                destination = BackupStorageDestination.DRIVE
            )
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
                                    color = if (hasSavedLocation) Color(0xFF2E7D32) else Color(0xFFF57F17),
                                    shape = RoundedCornerShape(50)
                                )
                        )
                        Text(
                            text = if (hasSavedLocation) "Ubicacion configurada" else "Sin ubicacion configurada",
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
                            "Usando Google Drive mediante selector de archivos del sistema."
                        },
                        style = MaterialTheme.typography.bodySmall,
                        color = MaterialTheme.colorScheme.onSurfaceVariant
                    )
                    if (!isLocalSelected) {
                        OutlinedButton(
                            onClick = {
                                if (hasSavedLocation) {
                                    launchDriveOpenFile(
                                        context = context,
                                        launcher = driveExistingFileLauncher,
                                        onError = viewModel::reportarError
                                    )
                                } else {
                                    launchDriveCreateFile(
                                        context = context,
                                        launcher = driveFileBackupLauncher,
                                        onError = viewModel::reportarError
                                    )
                                }
                            },
                            enabled = uiState.licenseActive,
                            modifier = Modifier.fillMaxWidth()
                        ) {
                            Text(if (hasSavedLocation) "Cambiar archivo de Drive" else "Conectar Drive")
                        }
                    }
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
                        enabled = uiState.licenseActive,
                        onClick = {
                            if (isLocalSelected) {
                                folderBackupLauncher.launch(Intent(Intent.ACTION_OPEN_DOCUMENT_TREE))
                            } else {
                                if (hasSavedLocation) {
                                    launchDriveOpenFile(
                                        context = context,
                                        launcher = driveExistingFileLauncher,
                                        onError = viewModel::reportarError
                                    )
                                } else {
                                    launchDriveCreateFile(
                                        context = context,
                                        launcher = driveFileBackupLauncher,
                                        onError = viewModel::reportarError
                                    )
                                }
                            }
                        }
                    ) { Text(if (isLocalSelected) "Configurar ubicacion" else "Conectar Drive") }

                    Button(
                        modifier = Modifier.fillMaxWidth(),
                        enabled = uiState.licenseActive,
                        onClick = {
                            if (hasSavedLocation) {
                                viewModel.generarRespaldoEnUbicacionGuardada(
                                    destination = storageDestination,
                                    onLocationMissing = {
                                        if (isLocalSelected) {
                                            folderBackupLauncher.launch(Intent(Intent.ACTION_OPEN_DOCUMENT_TREE))
                                        } else {
                                            if (hasSavedLocation) {
                                                launchDriveOpenFile(
                                                    context = context,
                                                    launcher = driveExistingFileLauncher,
                                                    onError = viewModel::reportarError
                                                )
                                            } else {
                                                launchDriveCreateFile(
                                                    context = context,
                                                    launcher = driveFileBackupLauncher,
                                                    onError = viewModel::reportarError
                                                )
                                            }
                                        }
                                    }
                                )
                            } else {
                                if (isLocalSelected) {
                                    folderBackupLauncher.launch(Intent(Intent.ACTION_OPEN_DOCUMENT_TREE))
                                } else {
                                    if (hasSavedLocation) {
                                        launchDriveOpenFile(
                                            context = context,
                                            launcher = driveExistingFileLauncher,
                                            onError = viewModel::reportarError
                                        )
                                    } else {
                                        launchDriveCreateFile(
                                            context = context,
                                            launcher = driveFileBackupLauncher,
                                            onError = viewModel::reportarError
                                        )
                                    }
                                }
                            }
                        }
                    ) { Text("Crear respaldo ahora") }

                    OutlinedButton(
                        modifier = Modifier.fillMaxWidth(),
                        enabled = uiState.licenseActive,
                        onClick = { restoreBackupLauncher.launch(arrayOf("application/json", "text/plain", "*/*")) }
                    ) { Text("Restaurar respaldo") }

                    OutlinedButton(
                        modifier = Modifier.fillMaxWidth(),
                        enabled = uiState.licenseActive && hasSavedLocation,
                        onClick = {
                            viewModel.getSavedBackupUri(storageDestination) { uri ->
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
                        enabled = uiState.licenseActive && hasSavedLocation,
                        onClick = {
                            viewModel.getSavedBackupUri(storageDestination) { uri ->
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

private fun launchDriveCreateFile(
    context: android.content.Context,
    launcher: androidx.activity.result.ActivityResultLauncher<Intent>,
    onError: (String) -> Unit
) {
    val intent = Intent(Intent.ACTION_CREATE_DOCUMENT).apply {
        addCategory(Intent.CATEGORY_OPENABLE)
        type = "application/json"
        setPackage("com.google.android.apps.docs")
        putExtra(Intent.EXTRA_TITLE, BackupManager.BACKUP_FILE_NAME)
        val driveRoot = DocumentsContract.buildRootUri("com.google.android.apps.docs.storage", "root")
        putExtra(DocumentsContract.EXTRA_INITIAL_URI, driveRoot)
    }

    if (intent.resolveActivity(context.packageManager) == null) {
        onError("No se pudo abrir Google Drive. Verifica la app de Drive.")
        return
    }
    runCatching { launcher.launch(intent) }
        .onFailure { onError("No se pudo abrir Google Drive") }
}

private fun launchDriveOpenFile(
    context: android.content.Context,
    launcher: androidx.activity.result.ActivityResultLauncher<Intent>,
    onError: (String) -> Unit
) {
    val intent = Intent(Intent.ACTION_OPEN_DOCUMENT).apply {
        addCategory(Intent.CATEGORY_OPENABLE)
        type = "application/json"
        putExtra(Intent.EXTRA_MIME_TYPES, arrayOf("application/json", "text/plain"))
        setPackage("com.google.android.apps.docs")
        val driveRoot = DocumentsContract.buildRootUri("com.google.android.apps.docs.storage", "root")
        putExtra(DocumentsContract.EXTRA_INITIAL_URI, driveRoot)
        addFlags(Intent.FLAG_GRANT_READ_URI_PERMISSION)
        addFlags(Intent.FLAG_GRANT_WRITE_URI_PERMISSION)
        addFlags(Intent.FLAG_GRANT_PERSISTABLE_URI_PERMISSION)
    }

    if (intent.resolveActivity(context.packageManager) == null) {
        onError("No se pudo abrir Google Drive. Verifica la app de Drive.")
        return
    }
    runCatching { launcher.launch(intent) }
        .onFailure { onError("No se pudo abrir Google Drive") }
}

private fun Long?.toDisplayDateTime(): String {
    if (this == null || this <= 0L) return "Nunca"
    val formatter = SimpleDateFormat("dd/MM/yyyy HH:mm", Locale.getDefault())
    return formatter.format(Date(this))
}
