package com.prestamos.app.ui.screen

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
import androidx.compose.runtime.getValue
import androidx.compose.runtime.mutableStateOf
import androidx.compose.runtime.remember
import androidx.compose.runtime.rememberCoroutineScope
import androidx.compose.runtime.setValue
import androidx.compose.ui.Modifier
import androidx.compose.ui.platform.LocalContext
import androidx.compose.ui.unit.dp
import com.prestamos.app.data.license.LicenseManager
import com.prestamos.app.data.local.DatabaseBackupManager
import kotlinx.coroutines.launch

@Composable
fun BackupScreen() {
    val context = LocalContext.current
    val scope = rememberCoroutineScope()
    val snackbarHostState = remember { SnackbarHostState() }
    val licenseManager = remember(context) { LicenseManager(context) }
    var pendingImportUri by remember { mutableStateOf<Uri?>(null) }

    val exportBackupLauncher = rememberLauncherForActivityResult(
        contract = ActivityResultContracts.CreateDocument("application/octet-stream")
    ) { uri ->
        if (uri == null) return@rememberLauncherForActivityResult
        scope.launch {
            val status = runCatching { licenseManager.evaluateStatus() }.getOrNull()
            if (status?.isValid != true) {
                snackbarHostState.showSnackbar("Licencia inválida o vencida. No se permite exportar backup.")
                return@launch
            }
            DatabaseBackupManager.exportDatabase(context, uri)
                .onSuccess { snackbarHostState.showSnackbar("Backup exportado correctamente") }
                .onFailure { snackbarHostState.showSnackbar(it.message ?: "No se pudo exportar backup") }
        }
    }

    val importBackupLauncher = rememberLauncherForActivityResult(
        contract = ActivityResultContracts.OpenDocument()
    ) { uri ->
        if (uri == null) return@rememberLauncherForActivityResult
        pendingImportUri = uri
    }

    Column(
        modifier = Modifier
            .fillMaxSize()
            .padding(16.dp),
        verticalArrangement = Arrangement.spacedBy(12.dp)
    ) {
        Text("Backup y restauración", style = MaterialTheme.typography.headlineSmall)
        Text("Exporta o restaura una copia local de la base de datos (.db).")

        Button(
            modifier = Modifier.fillMaxWidth(),
            onClick = {
                scope.launch {
                    val status = runCatching { licenseManager.evaluateStatus() }.getOrNull()
                    if (status?.isValid != true) {
                        snackbarHostState.showSnackbar("Licencia inválida o vencida. No se permite exportar backup.")
                        return@launch
                    }
                    exportBackupLauncher.launch(DatabaseBackupManager.buildBackupFileName())
                }
            }
        ) {
            Text("Exportar backup")
        }

        Button(
            modifier = Modifier.fillMaxWidth(),
            onClick = {
                scope.launch {
                    val status = runCatching { licenseManager.evaluateStatus() }.getOrNull()
                    if (status?.isValid != true) {
                        snackbarHostState.showSnackbar("Licencia inválida o vencida. No se permite restaurar backup.")
                        return@launch
                    }
                    importBackupLauncher.launch(arrayOf("application/octet-stream", "application/x-sqlite3", "*/*"))
                }
            }
        ) {
            Text("Importar backup")
        }

        SnackbarHost(hostState = snackbarHostState)
    }

    pendingImportUri?.let { uri ->
        AlertDialog(
            onDismissRequest = { pendingImportUri = null },
            title = { Text("Confirmar restauración") },
            text = { Text("Restaurar este backup reemplazará todos los datos actuales. ¿Desea continuar?") },
            confirmButton = {
                TextButton(onClick = {
                    pendingImportUri = null
                    scope.launch {
                        val status = runCatching { licenseManager.evaluateStatus() }.getOrNull()
                        if (status?.isValid != true) {
                            snackbarHostState.showSnackbar("Licencia inválida o vencida. No se permite restaurar backup.")
                            return@launch
                        }
                        DatabaseBackupManager.importDatabase(context, uri)
                            .onSuccess {
                                snackbarHostState.showSnackbar("Backup restaurado. Reiniciando app...")
                                DatabaseBackupManager.restartApplication(context)
                            }
                            .onFailure {
                                snackbarHostState.showSnackbar(it.message ?: "No se pudo restaurar backup") }
                    }
                }) {
                    Text("Restaurar")
                }
            },
            dismissButton = {
                TextButton(onClick = { pendingImportUri = null }) { Text("Cancelar") }
            }
        )
    }
}
