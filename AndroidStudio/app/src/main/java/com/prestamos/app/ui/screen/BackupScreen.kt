package com.prestamos.app.ui.screen

import android.content.Intent
import android.net.Uri
import androidx.activity.compose.rememberLauncherForActivityResult
import androidx.activity.result.contract.ActivityResultContracts
import androidx.compose.foundation.layout.Arrangement
import androidx.compose.foundation.layout.Column
import androidx.compose.foundation.layout.fillMaxSize
import androidx.compose.foundation.layout.fillMaxWidth
import androidx.compose.foundation.layout.padding
import androidx.compose.material3.AlertDialog
import androidx.compose.material3.Button
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
import androidx.compose.ui.Modifier
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
    var pendingImportUri by remember { mutableStateOf<Uri?>(null) }

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

    LaunchedEffect(mensaje) {
        val text = mensaje ?: return@LaunchedEffect
        snackbarHostState.showSnackbar(text)
        viewModel.limpiarMensaje()
    }

    Column(
        modifier = Modifier
            .fillMaxSize()
            .padding(16.dp),
        verticalArrangement = Arrangement.spacedBy(12.dp)
    ) {
        Text("Respaldo", style = MaterialTheme.typography.headlineSmall)
        Text("Archivo único: ${BackupManager.BACKUP_FILE_NAME}")
        Text("Último respaldo realizado: ${uiState.lastBackupTimestamp.toDisplayDateTime()}")
        Text("Estado: ${if (uiState.hasSavedLocation) "Ubicación configurada" else "Sin ubicación configurada"}")

        Button(
            modifier = Modifier.fillMaxWidth(),
            onClick = { folderBackupLauncher.launch(null) }
        ) {
            Text("Configurar ubicación")
        }

        Button(
            modifier = Modifier.fillMaxWidth(),
            onClick = {
                if (uiState.hasSavedLocation) {
                    viewModel.generarRespaldoEnUbicacionGuardada(
                        onLocationMissing = {
                            folderBackupLauncher.launch(null)
                        }
                    )
                } else {
                    folderBackupLauncher.launch(null)
                }
            }
        ) {
            Text("Generar respaldo")
        }

        Button(
            modifier = Modifier.fillMaxWidth(),
            onClick = { restoreBackupLauncher.launch(arrayOf("application/json", "text/plain", "*/*")) }
        ) {
            Text("Restaurar respaldo")
        }

        Button(
            modifier = Modifier.fillMaxWidth(),
            enabled = uiState.hasSavedLocation,
            onClick = {
                viewModel.getSavedBackupUri { uri ->
                    if (uri == null) {
                        viewModel.limpiarMensaje()
                        return@getSavedBackupUri
                    }
                    val sendIntent = Intent(Intent.ACTION_SEND).apply {
                        type = "application/json"
                        putExtra(Intent.EXTRA_STREAM, uri)
                        putExtra(Intent.EXTRA_SUBJECT, "Respaldo de préstamos")
                        addFlags(Intent.FLAG_GRANT_READ_URI_PERMISSION)
                    }
                    context.startActivity(Intent.createChooser(sendIntent, "Compartir respaldo"))
                }
            }
        ) {
            Text("Compartir respaldo")
        }

        SnackbarHost(hostState = snackbarHostState)
    }

    pendingImportUri?.let { uri ->
        AlertDialog(
            onDismissRequest = { pendingImportUri = null },
            title = { Text("Confirmar restauración") },
            text = { Text("Esta acción reemplazará todos los datos actuales. ¿Desea continuar?") },
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

private fun Long?.toDisplayDateTime(): String {
    if (this == null || this <= 0L) return "Nunca"
    val formatter = SimpleDateFormat("dd/MM/yyyy HH:mm", Locale.getDefault())
    return formatter.format(Date(this))
}
