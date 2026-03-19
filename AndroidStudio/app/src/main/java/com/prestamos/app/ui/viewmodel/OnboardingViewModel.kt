package com.prestamos.app.ui.viewmodel

import android.app.Application
import androidx.lifecycle.AndroidViewModel
import androidx.lifecycle.viewModelScope
import com.prestamos.app.data.config.InitialSetupPreferences
import kotlinx.coroutines.flow.MutableStateFlow
import kotlinx.coroutines.flow.StateFlow
import kotlinx.coroutines.launch

data class OnboardingUiState(
    val loading: Boolean = true,
    val showOnboarding: Boolean = true,
    val step: Int = 0,
    val businessName: String = "",
    val mainCurrencyCode: String? = null,
    val secondaryCurrencyCode: String? = null,
    val errorMessage: String? = null
)

class OnboardingViewModel(application: Application) : AndroidViewModel(application) {
    private val prefs = InitialSetupPreferences(application)

    private val _uiState = MutableStateFlow(OnboardingUiState())
    val uiState: StateFlow<OnboardingUiState> = _uiState

    init {
        viewModelScope.launch {
            val isFirstRun = prefs.isFirstRun()
            _uiState.value = _uiState.value.copy(
                loading = false,
                showOnboarding = isFirstRun,
                step = 0
            )
        }
    }

    fun comenzar() {
        _uiState.value = _uiState.value.copy(step = 1, errorMessage = null)
    }

    fun updateBusinessName(value: String) {
        _uiState.value = _uiState.value.copy(
            businessName = value,
            errorMessage = null
        )
    }

    fun selectMainCurrency(code: String) {
        val secondary = _uiState.value.secondaryCurrencyCode
        _uiState.value = _uiState.value.copy(
            mainCurrencyCode = code,
            secondaryCurrencyCode = if (secondary == code) null else secondary,
            errorMessage = null
        )
    }

    fun selectSecondaryCurrency(code: String?) {
        _uiState.value = _uiState.value.copy(
            secondaryCurrencyCode = if (code == _uiState.value.mainCurrencyCode) null else code,
            errorMessage = null
        )
    }

    fun finalizarConfiguracion() {
        val current = _uiState.value
        val businessName = current.businessName.trim()
        val mainCurrency = current.mainCurrencyCode

        if (businessName.isBlank()) {
            _uiState.value = current.copy(errorMessage = "Ingresa el nombre del negocio o prestamista")
            return
        }
        if (mainCurrency.isNullOrBlank()) {
            _uiState.value = current.copy(errorMessage = "Selecciona una moneda principal")
            return
        }

        viewModelScope.launch {
            prefs.saveInitialSetup(
                businessName = businessName,
                mainCurrencyCode = mainCurrency,
                secondaryCurrencyCode = current.secondaryCurrencyCode
            )
            _uiState.value = _uiState.value.copy(
                showOnboarding = false,
                errorMessage = null
            )
        }
    }
}
