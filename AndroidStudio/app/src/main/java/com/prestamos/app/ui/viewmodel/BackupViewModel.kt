package com.prestamos.app.ui.viewmodel

import android.app.Application
import android.content.Intent
import android.net.Uri
import androidx.lifecycle.AndroidViewModel
import androidx.lifecycle.viewModelScope
import com.google.android.gms.auth.api.signin.GoogleSignIn
import com.google.android.gms.auth.api.signin.GoogleSignInAccount
import com.google.android.gms.auth.api.signin.GoogleSignInOptions
import com.google.android.gms.common.api.ApiException
import com.google.api.services.drive.DriveScopes
import com.prestamos.app.data.backup.BackupManager
import com.prestamos.app.data.backup.BackupStorageDestination
import com.prestamos.app.data.backup.DriveBackupManager
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
    val driveConnected: Boolean = false,
    val driveAccountEmail: String? = null,
    val lastBackupTimestamp: Long? = null,
    val licenseActive: Boolean = false
)

class BackupViewModel(application: Application) : AndroidViewModel(application) {
    private val backupManager = BackupManager(application)
    private val driveBackupManager = DriveBackupManager(application, backupManager)
    private val licenseManager = LicenseManager(application)
    val mensaje = MutableStateFlow<String?>(null)

    private val hasSavedLocationLocal = MutableStateFlow(false)
    private val driveConnected = MutableStateFlow(false)
    private val driveAccountEmail = MutableStateFlow<String?>(null)
    private val licenseActive = MutableStateFlow(false)

    private val signInClient by lazy {
        val options = GoogleSignInOptions.Builder(GoogleSignInOptions.DEFAULT_SIGN_IN)
            .requestEmail()
            .requestScopes(com.google.android.gms.common.api.Scope(DriveScopes.DRIVE_FILE))
            .build()
        GoogleSignIn.getClient(getApplication(), options)
    }

    val uiState: StateFlow<BackupUiState> = combine(
        hasSavedLocationLocal,
        driveConnected,
        driveAccountEmail,
        backupManager.observeLastBackupTimestamp(),
        licenseActive
    ) { hasLocationLocal, isDriveConnected, driveEmail, lastTimestamp, isLicenseActive ->
        BackupUiState(
            hasSavedLocationLocal = hasLocationLocal,
            driveConnected = isDriveConnected,
            driveAccountEmail = driveEmail,
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
        refreshDriveConnection()
    }

    fun getDriveSignInIntent(): Intent = signInClient.signInIntent

    fun handleDriveSignInResult(data: Intent?) {
        runCatching {
            val account = GoogleSignIn.getSignedInAccountFromIntent(data).getResult(ApiException::class.java)
            requireNotNull(account) { "No se pudo conectar con Google Drive" }
            account
        }.onSuccess { account ->
            driveConnected.value = true
            driveAccountEmail.value = account.email
            mensaje.value = "Drive conectado: ${account.email ?: "Cuenta Google"}"
        }.onFailure {
            driveConnected.value = false
            driveAccountEmail.value = null
            mensaje.value = "No se pudo conectar con Google Drive"
        }
    }

    fun desconectarDrive() {
        signInClient.signOut().addOnCompleteListener {
            driveConnected.value = false
            driveAccountEmail.value = null
            mensaje.value = "Drive desconectado"
        }
    }

    fun cambiarCuentaDrive(onReady: (Intent) -> Unit) {
        signInClient.signOut().addOnCompleteListener {
            driveConnected.value = false
            driveAccountEmail.value = null
            onReady(signInClient.signInIntent)
        }
    }

    fun generarRespaldoDrive(startedFromUi: Boolean = true) {
        viewModelScope.launch {
            if (!ensureLicenseActive()) return@launch
            val account = requireDriveAccount() ?: return@launch
            if (startedFromUi) mensaje.value = "Generando respaldo en Drive..."
            driveBackupManager.subirRespaldo(account)
                .onSuccess {
                    refreshSavedLocation()
                    mensaje.value = "Respaldo en Drive creado correctamente"
                }
                .onFailure {
                    mensaje.value = it.message ?: "Error al crear respaldo en Drive"
                }
        }
    }

    fun restaurarRespaldoDrive(onSuccess: () -> Unit = {}) {
        viewModelScope.launch {
            if (!ensureLicenseActive()) return@launch
            val account = requireDriveAccount() ?: return@launch
            driveBackupManager.restaurarRespaldo(account)
                .onSuccess {
                    mensaje.value = "Restauracion desde Drive completada. Reiniciando app..."
                    onSuccess()
                }
                .onFailure {
                    mensaje.value = it.message ?: "Error al restaurar desde Drive"
                }
        }
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
            if (destination == BackupStorageDestination.DRIVE) {
                generarRespaldoDrive(startedFromUi = startedFromUi)
                return@launch
            }
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
            if (destination == BackupStorageDestination.DRIVE) {
                onResult(null)
            } else {
                onResult(backupManager.getSavedBackupUri(destination))
            }
        }
    }

    fun refreshSavedLocation() {
        viewModelScope.launch {
            hasSavedLocationLocal.value = backupManager.hasSavedLocation(BackupStorageDestination.LOCAL)
            refreshDriveConnection()
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

    private fun refreshDriveConnection() {
        val account = GoogleSignIn.getLastSignedInAccount(getApplication())
        val hasScope = account?.grantedScopes?.any { it.scopeUri == DriveScopes.DRIVE_FILE } == true
        driveConnected.value = account != null && hasScope
        driveAccountEmail.value = account?.email
    }

    private fun requireDriveAccount(): GoogleSignInAccount? {
        val account = GoogleSignIn.getLastSignedInAccount(getApplication())
        val hasScope = account?.grantedScopes?.any { it.scopeUri == DriveScopes.DRIVE_FILE } == true
        if (account == null || !hasScope) {
            driveConnected.value = false
            driveAccountEmail.value = null
            mensaje.value = "Reconecta tu cuenta de Drive para otorgar permisos"
            return null
        }
        return account
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
