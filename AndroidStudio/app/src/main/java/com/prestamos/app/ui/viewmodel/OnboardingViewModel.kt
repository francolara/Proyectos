package com.prestamos.app.ui.viewmodel

import android.app.Application
import androidx.lifecycle.AndroidViewModel
import androidx.lifecycle.viewModelScope
import com.prestamos.app.data.config.InitialSetupPreferences
import com.prestamos.app.data.local.AppDatabase
import com.prestamos.app.data.local.entity.TipoCobroEntity
import com.prestamos.app.data.local.entity.TipoPago
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
    val defaultInterest: String = "",
    val allowedPaymentTypes: Set<TipoPago> = TipoPago.entries.toSet(),
    val collectionTypes: List<String> = emptyList(),
    val errorMessage: String? = null
)

class OnboardingViewModel(application: Application) : AndroidViewModel(application) {
    private val prefs = InitialSetupPreferences(application)
    private val tipoCobroDao = AppDatabase.getInstance(application).tipoCobroDao()

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

    fun updateDefaultInterest(value: String) {
        _uiState.value = _uiState.value.copy(
            defaultInterest = value,
            errorMessage = null
        )
    }

    fun toggleAllowedPaymentType(type: TipoPago) {
        val current = _uiState.value.allowedPaymentTypes.toMutableSet()
        if (type in current) current.remove(type) else current.add(type)
        _uiState.value = _uiState.value.copy(
            allowedPaymentTypes = current,
            errorMessage = null
        )
    }

    fun addCollectionType(nombre: String) {
        val clean = nombre.trim()
        if (clean.isBlank()) return
        val exists = _uiState.value.collectionTypes.any { it.equals(clean, ignoreCase = true) }
        if (exists) {
            _uiState.value = _uiState.value.copy(errorMessage = "Ese tipo de cobro ya fue agregado")
            return
        }
        _uiState.value = _uiState.value.copy(
            collectionTypes = _uiState.value.collectionTypes + clean,
            errorMessage = null
        )
    }

    fun removeCollectionType(nombre: String) {
        if (_uiState.value.collectionTypes.size <= 1) {
            _uiState.value = _uiState.value.copy(
                errorMessage = "Debes mantener al menos un tipo de cobro"
            )
            return
        }
        _uiState.value = _uiState.value.copy(
            collectionTypes = _uiState.value.collectionTypes.filterNot { it.equals(nombre, ignoreCase = true) },
            errorMessage = null
        )
    }

    fun finalizarConfiguracion() {
        val current = _uiState.value
        val businessName = current.businessName.trim()
        val mainCurrency = current.mainCurrencyCode

        if (businessName.isBlank()) {
            _uiState.value = current.copy(errorMessage = "Ingresa el Nombre del Negocio")
            return
        }
        if (mainCurrency.isNullOrBlank()) {
            _uiState.value = current.copy(errorMessage = "Selecciona una moneda principal")
            return
        }
        if (current.allowedPaymentTypes.isEmpty()) {
            _uiState.value = current.copy(errorMessage = "Selecciona al menos un tipo de pago")
            return
        }
        if (current.collectionTypes.isEmpty()) {
            _uiState.value = current.copy(errorMessage = "Agrega al menos un tipo de cobro")
            return
        }

        viewModelScope.launch {
            prefs.saveInitialSetup(
                businessName = businessName,
                mainCurrencyCode = mainCurrency,
                secondaryCurrencyCode = current.secondaryCurrencyCode,
                defaultInterest = current.defaultInterest,
                allowedPaymentTypes = current.allowedPaymentTypes
            )

            val now = System.currentTimeMillis()
            current.collectionTypes.forEach { nombre ->
                if (tipoCobroDao.contarPorNombre(nombre) == 0) {
                    tipoCobroDao.insertar(
                        TipoCobroEntity(
                            nombre = nombre.trim(),
                            fechaRegistro = now,
                            fechaModificacion = now
                        )
                    )
                }
            }
            _uiState.value = _uiState.value.copy(
                showOnboarding = false,
                errorMessage = null
            )
        }
    }
}
