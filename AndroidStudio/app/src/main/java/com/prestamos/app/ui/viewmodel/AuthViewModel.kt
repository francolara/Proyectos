package com.prestamos.app.ui.viewmodel

import android.app.Application
import androidx.lifecycle.AndroidViewModel
import androidx.lifecycle.viewModelScope
import com.prestamos.app.data.backup.BackupManager
import com.prestamos.app.data.backup.BackupStorageDestination
import com.prestamos.app.data.backup.DriveBackupManager
import com.google.android.gms.auth.api.signin.GoogleSignIn
import com.google.api.services.drive.DriveScopes
import com.prestamos.app.data.license.LicenseManager
import com.prestamos.app.data.license.LicenseType
import com.prestamos.app.data.security.AuthRepository
import kotlinx.coroutines.flow.MutableStateFlow
import kotlinx.coroutines.flow.SharingStarted
import kotlinx.coroutines.flow.StateFlow
import kotlinx.coroutines.flow.combine
import kotlinx.coroutines.flow.stateIn
import kotlinx.coroutines.launch

sealed class AuthState {
    data object Loading : AuthState()
    data object NeedsPinSetup : AuthState()
    data object Locked : AuthState()
    data object Unlocked : AuthState()
}

class AuthViewModel(application: Application) : AndroidViewModel(application) {
    private val authRepository = AuthRepository(application)
    private val backupManager = BackupManager(application)
    private val driveBackupManager = DriveBackupManager(application, backupManager)
    private val licenseManager = LicenseManager(application)
    private val inicializado = MutableStateFlow(false)

    @Volatile
    private var cierreEnProgreso: Boolean = false

    init {
        viewModelScope.launch {
            authRepository.reiniciarSesionAlAbrirApp()
            inicializado.value = true
        }
    }

    val authState: StateFlow<AuthState> = combine(
        inicializado,
        authRepository.pinConfigured,
        authRepository.sessionUnlocked
    ) { listo, configured, unlocked ->
        when {
            !listo -> AuthState.Loading
            !configured -> AuthState.NeedsPinSetup
            unlocked -> AuthState.Unlocked
            else -> AuthState.Locked
        }
    }.stateIn(
        scope = viewModelScope,
        started = SharingStarted.WhileSubscribed(5000),
        initialValue = AuthState.Loading
    )

    val mensaje = MutableStateFlow<String?>(null)

    fun crearPin(pin: String, confirmPin: String) {
        viewModelScope.launch {
            runCatching {
                require(pin.length == 6 && pin.all { it.isDigit() }) { "El PIN debe tener 6 digitos" }
                require(pin == confirmPin) { "Los PIN no coinciden" }
                authRepository.configurarPin(pin)
            }.onFailure {
                mensaje.value = it.message ?: "No se pudo configurar el PIN"
            }
        }
    }

    fun ingresarPin(pin: String) {
        viewModelScope.launch {
            runCatching {
                require(pin.length == 6 && pin.all { it.isDigit() }) { "PIN invalido" }
                val ok = authRepository.validarPin(pin)
                require(ok) { "PIN incorrecto" }
            }.onFailure {
                mensaje.value = it.message ?: "No se pudo validar el PIN"
            }
        }
    }

    fun desbloquearConBiometria() {
        viewModelScope.launch {
            runCatching {
                authRepository.desbloquearSesion()
            }.onFailure {
                mensaje.value = it.message ?: "No se pudo validar biometria"
            }
        }
    }

    // Bloqueo de sesion por inactividad o pantalla bloqueada: sin respaldo.
    fun bloquearSesion() {
        viewModelScope.launch {
            authRepository.bloquearSesion()
        }
    }

    // Cierre de sesion manual desde menu: intenta respaldo automatico antes de bloquear.
    fun cerrarSesionConRespaldo() {
        viewModelScope.launch {
            if (cierreEnProgreso) return@launch
            cierreEnProgreso = true
            try {
                val status = licenseManager.evaluateStatus()
                val paidType = status.licenseType == LicenseType.MENSUAL ||
                    status.licenseType == LicenseType.ANUAL ||
                    status.licenseType == LicenseType.FULL

                if (status.isValid && status.isActivated && paidType) {
                    mensaje.value = "Generando respaldo..."
                    val localConfigured = backupManager.hasSavedLocation(BackupStorageDestination.LOCAL)
                    val driveAccount = GoogleSignIn.getLastSignedInAccount(getApplication())
                    val driveConnected = driveAccount != null &&
                        driveAccount.grantedScopes.any { it.scopeUri == DriveScopes.DRIVE_FILE }

                    when {
                        driveConnected -> {
                            driveBackupManager.subirRespaldo(driveAccount!!)
                                .onFailure { driveError ->
                                    if (localConfigured) {
                                        backupManager.exportBackupToSavedLocation(BackupStorageDestination.LOCAL)
                                            .onFailure { localError ->
                                                mensaje.value = sanitizeBackupError(
                                                    localError,
                                                    sanitizeBackupError(driveError, "Error al crear respaldo")
                                                )
                                            }
                                    } else {
                                        mensaje.value = sanitizeBackupError(driveError, "Error al crear respaldo")
                                    }
                                }
                        }
                        localConfigured -> {
                            backupManager.exportBackupToSavedLocation(BackupStorageDestination.LOCAL)
                                .onFailure {
                                    mensaje.value = sanitizeBackupError(it, "Error al crear respaldo")
                                }
                        }
                        else -> {
                            mensaje.value = "Primero configura una ubicacion de respaldo"
                        }
                    }
                } else {
                    mensaje.value = "Licencia activa requerida para respaldo y restauracion"
                }

                authRepository.bloquearSesion()
            } finally {
                cierreEnProgreso = false
            }
        }
    }

    fun limpiarMensaje() {
        mensaje.value = null
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
                "Clave de respaldo no configurada o invalida."
            "googleapis.com/drive" in normalized || normalized.trimStart().startsWith("{") ->
                fallback
            else -> raw
        }
    }
}
