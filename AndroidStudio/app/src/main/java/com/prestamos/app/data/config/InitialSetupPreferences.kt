package com.prestamos.app.data.config

import android.content.Context
import com.prestamos.app.data.local.entity.TipoPago

class InitialSetupPreferences(context: Context) {
    private val prefs = context.getSharedPreferences(PREFS_NAME, Context.MODE_PRIVATE)
    private val defaultFreePaymentTypes = setOf(TipoPago.SEMANAL, TipoPago.MENSUAL)

    fun isFirstRun(): Boolean = prefs.getBoolean(KEY_FIRST_RUN, true)
    fun getBusinessName(): String = prefs.getString(KEY_BUSINESS_NAME, "").orEmpty()
    fun getMainCurrencyCode(): String? = prefs.getString(KEY_MAIN_CURRENCY_CODE, null)
    fun getSecondaryCurrencyCode(): String? = prefs.getString(KEY_SECONDARY_CURRENCY_CODE, null)
    fun getDefaultInterest(): String = prefs.getString(KEY_DEFAULT_INTEREST, "").orEmpty()
    fun getAllowedPaymentTypes(): Set<TipoPago> {
        val raw = prefs.getStringSet(KEY_ALLOWED_PAYMENT_TYPES, null)
        if (raw.isNullOrEmpty()) return defaultFreePaymentTypes
        return raw.mapNotNull { value ->
            runCatching { TipoPago.valueOf(value) }.getOrNull()
        }.toSet().ifEmpty { defaultFreePaymentTypes }
    }

    fun saveInitialSetup(
        businessName: String,
        mainCurrencyCode: String,
        secondaryCurrencyCode: String?,
        defaultInterest: String,
        allowedPaymentTypes: Set<TipoPago>
    ) {
        prefs.edit()
            .putBoolean(KEY_FIRST_RUN, false)
            .putString(KEY_BUSINESS_NAME, businessName.trim())
            .putString(KEY_MAIN_CURRENCY_CODE, mainCurrencyCode)
            .putString(KEY_SECONDARY_CURRENCY_CODE, secondaryCurrencyCode)
            .putString(KEY_DEFAULT_INTEREST, defaultInterest.trim())
            .putStringSet(KEY_ALLOWED_PAYMENT_TYPES, allowedPaymentTypes.map { it.name }.toSet())
            .apply()
    }

    fun updateConfiguration(
        businessName: String,
        mainCurrencyCode: String,
        secondaryCurrencyCode: String?,
        defaultInterest: String = getDefaultInterest(),
        allowedPaymentTypes: Set<TipoPago> = getAllowedPaymentTypes()
    ) {
        prefs.edit()
            .putString(KEY_BUSINESS_NAME, businessName.trim())
            .putString(KEY_MAIN_CURRENCY_CODE, mainCurrencyCode)
            .putString(KEY_SECONDARY_CURRENCY_CODE, secondaryCurrencyCode)
            .putString(KEY_DEFAULT_INTEREST, defaultInterest.trim())
            .putStringSet(KEY_ALLOWED_PAYMENT_TYPES, allowedPaymentTypes.map { it.name }.toSet())
            .apply()
    }

    companion object {
        private const val PREFS_NAME = "initial_setup"
        const val KEY_FIRST_RUN = "first_run"
        const val KEY_BUSINESS_NAME = "business_name"
        const val KEY_MAIN_CURRENCY_CODE = "main_currency_code"
        const val KEY_SECONDARY_CURRENCY_CODE = "secondary_currency_code"
        const val KEY_DEFAULT_INTEREST = "default_interest"
        const val KEY_ALLOWED_PAYMENT_TYPES = "allowed_payment_types"
    }
}
