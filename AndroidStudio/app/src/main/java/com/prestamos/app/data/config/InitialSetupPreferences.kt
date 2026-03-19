package com.prestamos.app.data.config

import android.content.Context

class InitialSetupPreferences(context: Context) {
    private val prefs = context.getSharedPreferences(PREFS_NAME, Context.MODE_PRIVATE)

    fun isFirstRun(): Boolean = prefs.getBoolean(KEY_FIRST_RUN, true)
    fun getMainCurrencyCode(): String? = prefs.getString(KEY_MAIN_CURRENCY_CODE, null)
    fun getSecondaryCurrencyCode(): String? = prefs.getString(KEY_SECONDARY_CURRENCY_CODE, null)

    fun saveInitialSetup(
        businessName: String,
        mainCurrencyCode: String,
        secondaryCurrencyCode: String?
    ) {
        prefs.edit()
            .putBoolean(KEY_FIRST_RUN, false)
            .putString(KEY_BUSINESS_NAME, businessName.trim())
            .putString(KEY_MAIN_CURRENCY_CODE, mainCurrencyCode)
            .putString(KEY_SECONDARY_CURRENCY_CODE, secondaryCurrencyCode)
            .apply()
    }

    companion object {
        private const val PREFS_NAME = "initial_setup"
        const val KEY_FIRST_RUN = "first_run"
        const val KEY_BUSINESS_NAME = "business_name"
        const val KEY_MAIN_CURRENCY_CODE = "main_currency_code"
        const val KEY_SECONDARY_CURRENCY_CODE = "secondary_currency_code"
    }
}
