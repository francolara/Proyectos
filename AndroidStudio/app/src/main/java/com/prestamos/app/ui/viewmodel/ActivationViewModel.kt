package com.prestamos.app.ui.viewmodel

import android.app.Activity
import android.app.Application
import androidx.lifecycle.AndroidViewModel
import androidx.lifecycle.viewModelScope
import com.android.billingclient.api.AcknowledgePurchaseParams
import com.android.billingclient.api.BillingClient
import com.android.billingclient.api.BillingClient.BillingResponseCode
import com.android.billingclient.api.BillingClient.ProductType
import com.android.billingclient.api.BillingFlowParams
import com.android.billingclient.api.BillingResult
import com.android.billingclient.api.ProductDetails
import com.android.billingclient.api.Purchase
import com.android.billingclient.api.PurchasesResponseListener
import com.android.billingclient.api.QueryProductDetailsParams
import com.android.billingclient.api.QueryPurchasesParams
import com.prestamos.app.BuildConfig
import com.prestamos.app.data.license.LicenseManager
import com.prestamos.app.data.license.LicenseStatus
import com.prestamos.app.data.license.LicenseType
import kotlinx.coroutines.flow.MutableStateFlow
import kotlinx.coroutines.flow.StateFlow
import kotlinx.coroutines.launch

data class ActivationUiState(
    val loading: Boolean = true,
    val status: LicenseStatus = LicenseStatus(),
    val activationKey: String = "",
    val playBillingReady: Boolean = false,
    val playAvailablePlans: Set<LicenseType> = emptySet(),
    val playPlanPrices: Map<LicenseType, String> = emptyMap()
) {
    val canAccessApp: Boolean get() = !loading && status.isValid
}

class ActivationViewModel(application: Application) : AndroidViewModel(application) {
    private val licenseManager = LicenseManager(application)
    private var billingClient: BillingClient? = null
    private val productDetailsByType = mutableMapOf<LicenseType, ProductDetails>()

    private val productConfigByType: Map<LicenseType, PlayProductConfig> = mapOf(
        LicenseType.MENSUAL to PlayProductConfig("pro_mensual_sub", ProductType.SUBS),
        LicenseType.ANUAL to PlayProductConfig("pro_anual_sub", ProductType.SUBS),
        LicenseType.FULL to PlayProductConfig("pro_full_lifetime", ProductType.INAPP)
    )
    private val licenseByProductId = productConfigByType.entries.associate { (type, cfg) ->
        cfg.productId to type
    }

    private val _uiState = MutableStateFlow(ActivationUiState())
    val uiState: StateFlow<ActivationUiState> = _uiState

    val mensaje = MutableStateFlow<String?>(null)

    init {
        refreshStatus()
        if (BuildConfig.USE_PLAY_BILLING) {
            setupBilling()
        }
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
            if (BuildConfig.USE_PLAY_BILLING) {
                mensaje.value = "Esta version se activa desde Google Play"
                return@launch
            }
            val key = _uiState.value.activationKey
            runCatching {
                require(key.isNotBlank()) { "Ingresa la clave de activacion" }
                licenseManager.activate(key)
            }.onSuccess { status ->
                _uiState.value = _uiState.value.copy(status = status, activationKey = "")
                mensaje.value = "Licencia activada correctamente"
            }.onFailure {
                mensaje.value = it.message ?: "No se pudo activar la licencia"
            }
        }
    }

    fun startPlayPurchase(activity: Activity, type: LicenseType) {
        if (!BuildConfig.USE_PLAY_BILLING) {
            mensaje.value = "Esta compilacion no usa Google Play Billing"
            return
        }
        val billing = billingClient
        if (billing == null || !billing.isReady) {
            mensaje.value = "Google Play Billing no esta listo"
            return
        }

        val details = productDetailsByType[type]
        val config = productConfigByType[type]
        if (details == null || config == null) {
            mensaje.value = "Producto no configurado en Play Console para $type"
            return
        }

        val productParamsBuilder = BillingFlowParams.ProductDetailsParams.newBuilder()
            .setProductDetails(details)
        if (config.productType == ProductType.SUBS) {
            val offerToken = details.subscriptionOfferDetails
                ?.firstOrNull()
                ?.offerToken
                ?: run {
                    mensaje.value = "No se encontro oferta de suscripcion para $type"
                    return
                }
            productParamsBuilder.setOfferToken(offerToken)
        }

        val flowParams = BillingFlowParams.newBuilder()
            .setProductDetailsParamsList(listOf(productParamsBuilder.build()))
            .build()
        val result = billing.launchBillingFlow(activity, flowParams)
        if (result.responseCode != BillingResponseCode.OK) {
            mensaje.value = result.debugMessage.ifBlank { "No se pudo iniciar la compra" }
        }
    }

    fun limpiarMensaje() {
        mensaje.value = null
    }

    private fun setupBilling() {
        billingClient = BillingClient.newBuilder(getApplication())
            .setListener { billingResult, purchases ->
                handlePurchasesUpdated(billingResult, purchases)
            }
            .enablePendingPurchases()
            .build()
        connectBilling()
    }

    private fun connectBilling() {
        val billing = billingClient ?: return
        billing.startConnection(object : BillingClientStateSimpleListener {
            override fun onBillingSetupFinished(result: BillingResult) {
                if (result.responseCode == BillingResponseCode.OK) {
                    _uiState.value = _uiState.value.copy(playBillingReady = true)
                    queryPlayProducts()
                    restorePurchases()
                } else {
                    _uiState.value = _uiState.value.copy(playBillingReady = false)
                    mensaje.value = "No se pudo conectar con Google Play Billing"
                }
            }

            override fun onBillingServiceDisconnected() {
                _uiState.value = _uiState.value.copy(playBillingReady = false)
            }
        })
    }

    private fun queryPlayProducts() {
        val billing = billingClient ?: return
        val subsProducts = productConfigByType.values
            .filter { it.productType == ProductType.SUBS }
            .map { QueryProductDetailsParams.Product.newBuilder().setProductId(it.productId).setProductType(ProductType.SUBS).build() }
        val inAppProducts = productConfigByType.values
            .filter { it.productType == ProductType.INAPP }
            .map { QueryProductDetailsParams.Product.newBuilder().setProductId(it.productId).setProductType(ProductType.INAPP).build() }

        if (subsProducts.isNotEmpty()) {
            billing.queryProductDetailsAsync(
                QueryProductDetailsParams.newBuilder().setProductList(subsProducts).build()
            ) { _, productDetailsList ->
                mergeProductDetails(productDetailsList)
            }
        }
        if (inAppProducts.isNotEmpty()) {
            billing.queryProductDetailsAsync(
                QueryProductDetailsParams.newBuilder().setProductList(inAppProducts).build()
            ) { _, productDetailsList ->
                mergeProductDetails(productDetailsList)
            }
        }
    }

    private fun mergeProductDetails(items: List<ProductDetails>) {
        items.forEach { details ->
            val licenseType = licenseByProductId[details.productId] ?: return@forEach
            productDetailsByType[licenseType] = details
        }
        val prices = productDetailsByType.mapValues { (_, details) -> details.toDisplayPrice() }
        _uiState.value = _uiState.value.copy(
            playAvailablePlans = productDetailsByType.keys.toSet(),
            playPlanPrices = prices
        )
    }

    private fun restorePurchases() {
        val billing = billingClient ?: return
        billing.queryPurchasesAsync(
            QueryPurchasesParams.newBuilder().setProductType(ProductType.SUBS).build(),
            PurchasesResponseListener { _, purchases -> handleActivePurchases(purchases) }
        )
        billing.queryPurchasesAsync(
            QueryPurchasesParams.newBuilder().setProductType(ProductType.INAPP).build(),
            PurchasesResponseListener { _, purchases -> handleActivePurchases(purchases) }
        )
    }

    private fun handlePurchasesUpdated(
        billingResult: BillingResult,
        purchases: List<Purchase>?
    ) {
        when (billingResult.responseCode) {
            BillingResponseCode.OK -> handleActivePurchases(purchases.orEmpty())
            BillingResponseCode.USER_CANCELED -> mensaje.value = "Compra cancelada"
            else -> {
                if (billingResult.debugMessage.isNotBlank()) {
                    mensaje.value = billingResult.debugMessage
                }
            }
        }
    }

    private fun handleActivePurchases(purchases: List<Purchase>) {
        purchases.forEach { purchase ->
            when (purchase.purchaseState) {
                Purchase.PurchaseState.PURCHASED -> processPurchased(purchase)
                Purchase.PurchaseState.PENDING -> mensaje.value = "Compra pendiente de confirmacion"
                else -> Unit
            }
        }
    }

    private fun processPurchased(purchase: Purchase) {
        val plan = purchase.products
            .firstNotNullOfOrNull { productId -> licenseByProductId[productId] }
            ?: return

        viewModelScope.launch {
            runCatching {
                licenseManager.activateFromPlay(plan, purchase.purchaseToken)
            }.onSuccess { status ->
                _uiState.value = _uiState.value.copy(status = status)
                mensaje.value = "Licencia activada correctamente"
            }.onFailure {
                mensaje.value = it.message ?: "No se pudo activar la licencia"
            }
        }

        if (!purchase.isAcknowledged) {
            billingClient?.acknowledgePurchase(
                AcknowledgePurchaseParams.newBuilder()
                    .setPurchaseToken(purchase.purchaseToken)
                    .build()
            ) { }
        }
    }

    override fun onCleared() {
        billingClient?.endConnection()
        billingClient = null
        super.onCleared()
    }

    private fun ProductDetails.toDisplayPrice(): String {
        val subs = subscriptionOfferDetails
            ?.firstOrNull()
            ?.pricingPhases
            ?.pricingPhaseList
            ?.firstOrNull()
            ?.formattedPrice
        val inApp = oneTimePurchaseOfferDetails?.formattedPrice
        return subs ?: inApp ?: "No disponible"
    }

    private data class PlayProductConfig(
        val productId: String,
        val productType: String
    )
}

private interface BillingClientStateSimpleListener : com.android.billingclient.api.BillingClientStateListener
