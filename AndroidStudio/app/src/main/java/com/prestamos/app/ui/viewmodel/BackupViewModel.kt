package com.prestamos.app.ui.viewmodel

import android.app.Application
import android.content.Intent
import android.net.Uri
import androidx.lifecycle.AndroidViewModel
import androidx.lifecycle.viewModelScope
import com.prestamos.app.data.backup.BackupManager
import kotlinx.coroutines.flow.MutableStateFlow
import kotlinx.coroutines.flow.SharingStarted
import kotlinx.coroutines.flow.StateFlow
import kotlinx.coroutines.flow.combine
import kotlinx.coroutines.flow.stateIn
import kotlinx.coroutines.launch

data class BackupUiState(
    val hasSavedLocation: Boolean = false,
    val lastBackupTimestamp: Long? = null
)

class BackupViewModel(application: Application) : AndroidViewModel(application) {
    private val backupManager = BackupManager(application)
    val mensaje = MutableStateFlow<String?>(null)

    private val hasSavedLocation = MutableStateFlow(false)

    val uiState: StateFlow<BackupUiState> = combine(
        hasSavedLocation,
        backupManager.observeLastBackupTimestamp()
    ) { hasLocation, lastTimestamp ->
        BackupUiState(
            hasSavedLocation = hasLocation,
            lastBackupTimestamp = lastTimestamp
        )
    }.stateIn(
        scope = viewModelScope,
        started = SharingStarted.WhileSubscribed(5000),
        initialValue = BackupUiState()
    )

    init {
        refreshSavedLocation()
    }

    fun generarRespaldoEnUri(uri: Uri, persistPermission: Boolean, startedFromUi: Boolean = true) {
        viewModelScope.launch {
            if (startedFromUi) mensaje.value = "Generando respaldo..."
            backupManager.exportBackup(uri, persistPermission = persistPermission)
                .onSuccess {
                    refreshSavedLocation()
                    mensaje.value = "Respaldo creado correctamente"
                }
                .onFailure {
                    mensaje.value = it.message ?: "Error al crear respaldo"
                }
        }
    }

    fun generarRespaldoEnUbicacionGuardada(startedFromUi: Boolean = true, onLocationMissing: () -> Unit = {}) {
        viewModelScope.launch {
            if (startedFromUi) mensaje.value = "Generando respaldo..."
            backupManager.exportBackupToSavedLocation()
                .onSuccess {
                    refreshSavedLocation()
                    mensaje.value = "Respaldo creado correctamente"
                }
                .onFailure {
                    if (it is IllegalStateException) {
                        onLocationMissing()
                    } else {
                        mensaje.value = it.message ?: "Error al crear respaldo"
                    }
                }
        }
    }

    fun restaurarRespaldo(uri: Uri, onSuccess: () -> Unit = {}) {
        viewModelScope.launch {
            backupManager.importBackup(uri)
                .onSuccess {
                    mensaje.value = "Restauración completada. Reiniciando app..."
                    onSuccess()
                }
                .onFailure { mensaje.value = it.message ?: "Error al restaurar respaldo" }
        }
    }

    fun generarRespaldoEnCarpeta(uri: Uri, persistPermission: Boolean, startedFromUi: Boolean = true) {
        viewModelScope.launch {
            if (startedFromUi) mensaje.value = "Generando respaldo..."
            backupManager.exportBackupToFolder(uri, persistPermission = persistPermission)
                .onSuccess {
                    refreshSavedLocation()
                    mensaje.value = "Respaldo creado correctamente"
                }
                .onFailure {
                    mensaje.value = it.message ?: "Error al crear respaldo"
                }
        }
    }

    fun getSavedBackupUri(onResult: (Uri?) -> Unit) {
        viewModelScope.launch {
            onResult(backupManager.getSavedBackupUri())
        }
    }

    fun refreshSavedLocation() {
        viewModelScope.launch {
            hasSavedLocation.value = backupManager.hasSavedLocation()
        }
    }

    fun persistReadPermission(uri: Uri) {
        runCatching {
            getApplication<Application>().contentResolver.takePersistableUriPermission(
                uri,
                Intent.FLAG_GRANT_READ_URI_PERMISSION
            )
        }
    }

    fun limpiarMensaje() {
        mensaje.value = null
    }
}
