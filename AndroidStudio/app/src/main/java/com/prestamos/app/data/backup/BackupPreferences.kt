package com.prestamos.app.data.backup

import android.content.Context
import androidx.datastore.preferences.core.MutablePreferences
import androidx.datastore.preferences.core.longPreferencesKey
import androidx.datastore.preferences.core.stringPreferencesKey
import androidx.datastore.preferences.core.edit
import androidx.datastore.preferences.preferencesDataStore
import kotlinx.coroutines.flow.Flow
import kotlinx.coroutines.flow.map

private val Context.backupDataStore by preferencesDataStore(name = "backup_preferences")

class BackupPreferences(private val context: Context) {
    private val backupUriLegacyKey = stringPreferencesKey("backup_uri")
    private val backupUriLocalKey = stringPreferencesKey("backup_uri_local")
    private val backupUriDriveKey = stringPreferencesKey("backup_uri_drive")
    private val lastBackupTimestampKey = longPreferencesKey("last_backup_timestamp")
    private val backupPasswordKey = stringPreferencesKey("backup_password")

    val lastBackupTimestamp: Flow<Long?> = context.backupDataStore.data.map { it[lastBackupTimestampKey] }
    val backupPassword: Flow<String?> = context.backupDataStore.data.map { it[backupPasswordKey] }

    fun backupUri(destination: BackupStorageDestination): Flow<String?> = context.backupDataStore.data.map { prefs ->
        when (destination) {
            BackupStorageDestination.LOCAL -> prefs[backupUriLocalKey] ?: prefs[backupUriLegacyKey]
            BackupStorageDestination.DRIVE -> prefs[backupUriDriveKey]
        }
    }

    suspend fun saveBackupUri(uri: String, destination: BackupStorageDestination) {
        context.backupDataStore.edit { prefs: MutablePreferences ->
            when (destination) {
                BackupStorageDestination.LOCAL -> {
                    prefs[backupUriLocalKey] = uri
                    prefs[backupUriLegacyKey] = uri
                }
                BackupStorageDestination.DRIVE -> prefs[backupUriDriveKey] = uri
            }
        }
    }

    suspend fun saveLastBackupTimestamp(timestamp: Long) {
        context.backupDataStore.edit { prefs: MutablePreferences ->
            prefs[lastBackupTimestampKey] = timestamp
        }
    }

    suspend fun clearBackupUri(destination: BackupStorageDestination) {
        context.backupDataStore.edit { prefs: MutablePreferences ->
            when (destination) {
                BackupStorageDestination.LOCAL -> {
                    prefs.remove(backupUriLocalKey)
                    prefs.remove(backupUriLegacyKey)
                }
                BackupStorageDestination.DRIVE -> prefs.remove(backupUriDriveKey)
            }
        }
    }

    suspend fun saveBackupPassword(password: String) {
        context.backupDataStore.edit { prefs: MutablePreferences ->
            prefs[backupPasswordKey] = password
        }
    }

    suspend fun clearBackupPassword() {
        context.backupDataStore.edit { prefs: MutablePreferences ->
            prefs.remove(backupPasswordKey)
        }
    }
}
