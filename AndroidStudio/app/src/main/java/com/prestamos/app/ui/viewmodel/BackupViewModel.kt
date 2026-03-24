package com.prestamos.app.ui.viewmodel

import android.app.Application
import android.content.Intent
import android.net.Uri
import androidx.lifecycle.AndroidViewModel
import androidx.lifecycle.viewModelScope
import com.prestamos.app.data.backup.BackupManager
import com.prestamos.app.data.backup.BackupStorageDestination
import com.prestamos.app.data.license.LicenseManager
import com.prestamos.app.data.license.LicenseType
import kotlinx.coroutines.flow.MutableStateFlow
import kotlinx.coroutines.flow.SharingStarted
import kotlinx.coroutines.flow.StateFlow
import kotlinx.coroutines.flow.combine
import kotlinx.coroutines.flow.stateIn
import kotlinx.coroutines.launch

data class BackupUiState(
    val hasSavedLocationLocal: Boolean = false,
    val hasSavedLocationDrive: Boolean = false,
    val lastBackupTimestamp: Long? = null,
    val licenseActive: Boolean = false
)

class BackupViewModel(application: Application) : AndroidViewModel(application) {
    private val backupManager = BackupManager(application)
    private val licenseManager = LicenseManager(application)
    val mensaje = MutableStateFlow<String?>(null)

    private val hasSavedLocationLocal = MutableStateFlow(false)
    private val hasSavedLocationDrive = MutableStateFlow(false)
    private val licenseActive = MutableStateFlow(false)

    val uiState: StateFlow<BackupUiState> = combine(
        hasSavedLocationLocal,
        hasSavedLocationDrive,
        backupManager.observeLastBackupTimestamp(),
        licenseActive
    ) { hasLocationLocal, hasLocationDrive, lastTimestamp, isLicenseActive ->
        BackupUiState(
            hasSavedLocationLocal = hasLocationLocal,
            hasSavedLocationDrive = hasLocationDrive,
            lastBackupTimestamp = lastTimestamp,
            licenseActive = isLicenseActive
        )
    }.stateIn(
        scope = viewModelScope,
        started = SharingStarted.WhileSubscribed(5000),
        initialValue = BackupUiState()
    )

    init {
        refreshSavedLocation()
        refreshLicenseStatus()
    }

    fun generarRespaldoEnUri(
        uri: Uri,
        persistPermission: Boolean,
        destination: BackupStorageDestination,
        startedFromUi: Boolean = true
    ) {
        viewModelScope.launch {
            if (!ensureLicenseActive()) return@launch
            if (startedFromUi) mensaje.value = "Generando respaldo..."
            backupManager.exportBackup(uri, persistPermission = persistPermission, destination = destination)
                .onSuccess {
                    refreshSavedLocation()
                    mensaje.value = "Respaldo creado correctamente"
                }
                .onFailure {
                    mensaje.value = it.message ?: "Error al crear respaldo"
                }
        }
    }

    fun generarRespaldoEnUbicacionGuardada(
        destination: BackupStorageDestination,
        startedFromUi: Boolean = true,
        onLocationMissing: () -> Unit = {}
    ) {
        viewModelScope.launch {
            if (!ensureLicenseActive()) return@launch
            if (startedFromUi) mensaje.value = "Generando respaldo..."
            backupManager.exportBackupToSavedLocation(destination)
                .onSuccess {
                    refreshSavedLocation()
                    mensaje.value = "Respaldo creado correctamente"
                }
                .onFailure {
                    if (it is IllegalStateException) {
                        refreshSavedLocation()
                        mensaje.value = it.message ?: "Ubicacion de respaldo no valida"
                        onLocationMissing()
                    } else {
                        mensaje.value = it.message ?: "Error al crear respaldo"
                    }
                }
        }
    }

    fun restaurarRespaldo(uri: Uri, onSuccess: () -> Unit = {}) {
        viewModelScope.launch {
            if (!ensureLicenseActive()) return@launch
            backupManager.importBackup(uri)
                .onSuccess {
                    mensaje.value = "Restauracion completada. Reiniciando app..."
                    onSuccess()
                }
                .onFailure { mensaje.value = it.message ?: "Error al restaurar respaldo" }
        }
    }

    fun generarRespaldoEnCarpeta(
        uri: Uri,
        persistPermission: Boolean,
        destination: BackupStorageDestination,
        startedFromUi: Boolean = true
    ) {
        viewModelScope.launch {
            if (!ensureLicenseActive()) return@launch
            if (startedFromUi) mensaje.value = "Generando respaldo..."
            backupManager.exportBackupToFolder(uri, persistPermission = persistPermission, destination = destination)
                .onSuccess {
                    refreshSavedLocation()
                    mensaje.value = "Respaldo creado correctamente"
                }
                .onFailure {
                    mensaje.value = it.message ?: "Error al crear respaldo"
                }
        }
    }

    fun getSavedBackupUri(destination: BackupStorageDestination, onResult: (Uri?) -> Unit) {
        viewModelScope.launch {
            onResult(backupManager.getSavedBackupUri(destination))
        }
    }

    fun refreshSavedLocation() {
        viewModelScope.launch {
            hasSavedLocationLocal.value = backupManager.hasSavedLocation(BackupStorageDestination.LOCAL)
            hasSavedLocationDrive.value = backupManager.hasSavedLocation(BackupStorageDestination.DRIVE)
        }
    }

    fun refreshLicenseStatus() {
        viewModelScope.launch {
            licenseActive.value = isPaidLicenseActive()
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

    fun reportarError(texto: String) {
        mensaje.value = texto
    }

    private suspend fun ensureLicenseActive(): Boolean {
        val active = isPaidLicenseActive()
        licenseActive.value = active
        if (!active) {
            mensaje.value = "Licencia activa requerida para respaldo y restauracion"
        }
        return active
    }

    private suspend fun isPaidLicenseActive(): Boolean {
        val status = licenseManager.evaluateStatus()
        val paidType = status.licenseType == LicenseType.MENSUAL ||
            status.licenseType == LicenseType.ANUAL ||
            status.licenseType == LicenseType.FULL
        return status.isValid && status.isActivated && paidType
    }
}
