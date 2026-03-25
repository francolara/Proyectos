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
    val licenseActive: Boolean = false,
    val hasBackupPassword: Boolean = false
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

    private val baseUiState = combine(
        hasSavedLocationLocal,
        driveConnected,
        driveAccountEmail,
        backupManager.observeLastBackupTimestamp()
    ) { hasLocationLocal, isDriveConnected, driveEmail, lastTimestamp ->
        BackupUiState(
            hasSavedLocationLocal = hasLocationLocal,
            driveConnected = isDriveConnected,
            driveAccountEmail = driveEmail,
            lastBackupTimestamp = lastTimestamp
        )
    }

    val uiState: StateFlow<BackupUiState> = combine(
        baseUiState,
        licenseActive,
        backupManager.observeHasBackupPassword()
    ) { base, isLicenseActive, hasPassword ->
        base.copy(
            licenseActive = isLicenseActive,
            hasBackupPassword = hasPassword
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
                    mensaje.value = sanitizeBackupError(it, "Error al crear respaldo en Drive")
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
                    mensaje.value = sanitizeBackupError(it, "Error al restaurar desde Drive")
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
                    mensaje.value = sanitizeBackupError(it, "Error al crear respaldo")
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
                        mensaje.value = sanitizeBackupError(it, "Ubicacion de respaldo no valida")
                        onLocationMissing()
                    } else {
                        mensaje.value = sanitizeBackupError(it, "Error al crear respaldo")
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
                .onFailure { mensaje.value = sanitizeBackupError(it, "Error al restaurar respaldo") }
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
                    mensaje.value = sanitizeBackupError(it, "Error al crear respaldo")
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

    fun guardarClaveRespaldo(password: String, confirmPassword: String) {
        viewModelScope.launch {
            if (password != confirmPassword) {
                mensaje.value = "Las claves no coinciden"
                return@launch
            }
            runCatching {
                backupManager.setBackupPassword(password.trim())
            }.onSuccess {
                mensaje.value = "Clave de respaldo guardada"
            }.onFailure {
                mensaje.value = it.message ?: "No se pudo guardar la clave de respaldo"
            }
        }
    }

    fun eliminarClaveRespaldo() {
        viewModelScope.launch {
            backupManager.clearBackupPassword()
            mensaje.value = "Clave de respaldo eliminada"
        }
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

    private fun sanitizeBackupError(error: Throwable, fallback: String): String {
        val raw = error.message.orEmpty()
        val normalized = raw.lowercase()
        return when {
            raw.isBlank() -> fallback
            "insufficientscopes" in normalized || "403" in normalized ->
                "Permisos de Drive insuficientes. Reconecta tu cuenta de Drive."
            "access_denied" in normalized || "acceso bloqueado" in normalized ->
                "No se pudo autorizar Drive con esta cuenta."
            "unable to resolve host" in normalized || "network" in normalized || "timeout" in normalized ->
                "No hay conexion estable. Intenta nuevamente."
            "clave de respaldo" in normalized || "tag mismatch" in normalized ->
                "Clave de respaldo incorrecta o respaldo invalido."
            "googleapis.com/drive" in normalized || normalized.trimStart().startsWith("{") ->
                fallback
            else -> raw
        }
    }
}
