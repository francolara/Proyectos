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
    private val backupUriKey = stringPreferencesKey("backup_uri")
    private val lastBackupTimestampKey = longPreferencesKey("last_backup_timestamp")

    val backupUri: Flow<String?> = context.backupDataStore.data.map { it[backupUriKey] }
    val lastBackupTimestamp: Flow<Long?> = context.backupDataStore.data.map { it[lastBackupTimestampKey] }

    suspend fun saveBackupUri(uri: String) {
        context.backupDataStore.edit { prefs: MutablePreferences ->
            prefs[backupUriKey] = uri
        }
    }

    suspend fun saveLastBackupTimestamp(timestamp: Long) {
        context.backupDataStore.edit { prefs: MutablePreferences ->
            prefs[lastBackupTimestampKey] = timestamp
        }
    }

    suspend fun clearBackupUri() {
        context.backupDataStore.edit { prefs: MutablePreferences ->
            prefs.remove(backupUriKey)
        }
    }
}
