package com.prestamos.app.ui.viewmodel

import android.app.Application
import androidx.lifecycle.AndroidViewModel
import androidx.lifecycle.viewModelScope
import com.prestamos.app.data.license.LicenseManager
import com.prestamos.app.data.license.LicenseStatus
import kotlinx.coroutines.flow.MutableStateFlow
import kotlinx.coroutines.flow.StateFlow
import kotlinx.coroutines.launch

data class ActivationUiState(
    val loading: Boolean = true,
    val status: LicenseStatus = LicenseStatus(),
    val activationKey: String = ""
) {
    val canAccessApp: Boolean get() = !loading && status.isValid
}

class ActivationViewModel(application: Application) : AndroidViewModel(application) {
    private val licenseManager = LicenseManager(application)

    private val _uiState = MutableStateFlow(ActivationUiState())
    val uiState: StateFlow<ActivationUiState> = _uiState

    val mensaje = MutableStateFlow<String?>(null)

    init {
        refreshStatus()
    }

    fun onActivationKeyChanged(value: String) {
        _uiState.value = _uiState.value.copy(activationKey = value.trim())
    }

    fun refreshStatus() {
        viewModelScope.launch {
            _uiState.value = _uiState.value.copy(loading = true)
            runCatching {
                licenseManager.evaluateStatus()
            }.onSuccess { status ->
                _uiState.value = _uiState.value.copy(loading = false, status = status)
            }.onFailure {
                _uiState.value = _uiState.value.copy(loading = false)
                mensaje.value = it.message ?: "No se pudo validar la licencia"
            }
        }
    }

    fun activate() {
        viewModelScope.launch {
            val key = _uiState.value.activationKey
            runCatching {
                require(key.isNotBlank()) { "Ingresa la clave de activaciÃ³n" }
                licenseManager.activate(key)
            }.onSuccess { status ->
                _uiState.value = _uiState.value.copy(status = status, activationKey = "")
                mensaje.value = "Licencia activada correctamente"
            }.onFailure {
                mensaje.value = it.message ?: "No se pudo activar la licencia"
            }
        }
    }

    fun limpiarMensaje() {
        mensaje.value = null
    }
}

