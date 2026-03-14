package com.prestamos.app.data.license

import android.content.Context
import androidx.datastore.preferences.core.MutablePreferences
import androidx.datastore.preferences.core.booleanPreferencesKey
import androidx.datastore.preferences.core.longPreferencesKey
import androidx.datastore.preferences.core.stringPreferencesKey
import androidx.datastore.preferences.core.edit
import androidx.datastore.preferences.preferencesDataStore
import kotlinx.coroutines.flow.Flow
import kotlinx.coroutines.flow.map

private val Context.licenseDataStore by preferencesDataStore(name = "license_preferences")

class LicensePreferences(private val context: Context) {
    private val deviceCodeKey = stringPreferencesKey("device_code")
    private val licenseTypeKey = stringPreferencesKey("license_type")
    private val isActivatedKey = booleanPreferencesKey("is_activated")
    private val trialStartDateKey = longPreferencesKey("trial_start_date")
    private val trialEndDateKey = longPreferencesKey("trial_end_date")
    private val activationDateKey = longPreferencesKey("activation_date")
    private val expirationDateKey = longPreferencesKey("expiration_date")
    private val licenseKeyKey = stringPreferencesKey("license_key")
    private val lastValidUseDateKey = longPreferencesKey("last_valid_use_date")

    val rawData: Flow<Map<String, Any?>> = context.licenseDataStore.data.map { prefs ->
        mapOf(
            "deviceCode" to prefs[deviceCodeKey],
            "licenseType" to prefs[licenseTypeKey],
            "isActivated" to prefs[isActivatedKey],
            "trialStartDate" to prefs[trialStartDateKey],
            "trialEndDate" to prefs[trialEndDateKey],
            "activationDate" to prefs[activationDateKey],
            "expirationDate" to prefs[expirationDateKey],
            "licenseKey" to prefs[licenseKeyKey],
            "lastValidUseDate" to prefs[lastValidUseDateKey]
        )
    }

    suspend fun initializeIfNeeded(deviceCode: String, now: Long, trialEnd: Long) {
        context.licenseDataStore.edit { prefs: MutablePreferences ->
            if (prefs[deviceCodeKey] == null) prefs[deviceCodeKey] = deviceCode
            if (prefs[trialStartDateKey] == null) prefs[trialStartDateKey] = now
            if (prefs[trialEndDateKey] == null) prefs[trialEndDateKey] = trialEnd
            if (prefs[licenseTypeKey] == null) prefs[licenseTypeKey] = LicenseType.TRIAL.name
            if (prefs[isActivatedKey] == null) prefs[isActivatedKey] = false
            if (prefs[lastValidUseDateKey] == null) prefs[lastValidUseDateKey] = now
        }
    }

    suspend fun saveLastValidUseDate(now: Long) {
        context.licenseDataStore.edit { prefs ->
            prefs[lastValidUseDateKey] = now
        }
    }

    suspend fun saveActivation(
        type: LicenseType,
        licenseKey: String,
        activationDate: Long,
        expirationDate: Long?
    ) {
        context.licenseDataStore.edit { prefs ->
            prefs[licenseTypeKey] = type.name
            prefs[isActivatedKey] = true
            prefs[activationDateKey] = activationDate
            if (expirationDate != null) {
                prefs[expirationDateKey] = expirationDate
            } else {
                prefs.remove(expirationDateKey)
            }
            prefs[licenseKeyKey] = licenseKey
            prefs[lastValidUseDateKey] = activationDate
        }
    }
}
