package com.prestamos.app.data.security

import android.content.Context
import androidx.datastore.preferences.core.MutablePreferences
import androidx.datastore.preferences.core.booleanPreferencesKey
import androidx.datastore.preferences.core.edit
import androidx.datastore.preferences.core.stringPreferencesKey
import androidx.datastore.preferences.preferencesDataStore
import kotlinx.coroutines.flow.Flow
import kotlinx.coroutines.flow.map

private val Context.dataStore by preferencesDataStore(name = "auth_preferences")

class AuthPreferences(private val context: Context) {
    private val pinHashKey = stringPreferencesKey("pin_hash")
    private val pinConfiguredKey = booleanPreferencesKey("pin_configured")
    private val sessionUnlockedKey = booleanPreferencesKey("session_unlocked")

    val pinHash: Flow<String?> = context.dataStore.data.map { it[pinHashKey] }
    val pinConfigured: Flow<Boolean> = context.dataStore.data.map { it[pinConfiguredKey] ?: false }
    val sessionUnlocked: Flow<Boolean> = context.dataStore.data.map { it[sessionUnlockedKey] ?: false }

    suspend fun savePinHash(hash: String) {
        context.dataStore.edit { prefs: MutablePreferences ->
            prefs[pinHashKey] = hash
            prefs[pinConfiguredKey] = true
        }
    }

    suspend fun setSessionUnlocked(unlocked: Boolean) {
        context.dataStore.edit { prefs: MutablePreferences ->
            prefs[sessionUnlockedKey] = unlocked
        }
    }
}
