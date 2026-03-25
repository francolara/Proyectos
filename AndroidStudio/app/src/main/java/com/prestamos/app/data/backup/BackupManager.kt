package com.prestamos.app.data.backup

import android.content.Context
import android.content.Intent
import android.net.Uri
import android.provider.DocumentsContract
import androidx.documentfile.provider.DocumentFile
import com.prestamos.app.data.local.DatabaseBackupManager
import java.io.FileNotFoundException
import java.security.GeneralSecurityException
import kotlinx.coroutines.Dispatchers
import kotlinx.coroutines.flow.first
import kotlinx.coroutines.flow.map
import kotlinx.coroutines.withContext

class BackupManager(private val context: Context) {
    private val prefs = BackupPreferences(context)

    suspend fun exportBackup(
        targetUri: Uri,
        persistPermission: Boolean,
        updateSavedUri: Boolean = true,
        destination: BackupStorageDestination = BackupStorageDestination.LOCAL
    ): Result<Unit> = withContext(Dispatchers.IO) {
        runCatching {
            if (destination == BackupStorageDestination.DRIVE && !isDriveUri(targetUri)) {
                throw IllegalStateException("Selecciona un archivo de Google Drive")
            }
            if (persistPermission) persistUriPermissions(targetUri)
            val encryptedBytes = exportBackupBytes().getOrThrow()
            context.contentResolver.openOutputStream(targetUri)?.use { output ->
                output.write(encryptedBytes)
            } ?: error("No se pudo abrir destino de exportacion")
            if (updateSavedUri) {
                prefs.saveBackupUri(targetUri.toString(), destination)
            }
            prefs.saveLastBackupTimestamp(System.currentTimeMillis())
        }
    }

    suspend fun exportBackupToFolder(
        folderUri: Uri,
        persistPermission: Boolean,
        destination: BackupStorageDestination = BackupStorageDestination.LOCAL
    ): Result<Unit> = withContext(Dispatchers.IO) {
        runCatching {
            if (destination == BackupStorageDestination.DRIVE && !isDriveUri(folderUri)) {
                throw IllegalStateException("Selecciona una carpeta de Google Drive")
            }
            if (persistPermission) persistUriPermissions(folderUri)
            val fileUri = createOrFindBackupFileInFolder(folderUri)
            exportBackup(fileUri, persistPermission = false, updateSavedUri = false, destination = destination).getOrThrow()
            prefs.saveBackupUri(folderUri.toString(), destination)
        }
    }

    suspend fun exportBackupToSavedLocation(
        destination: BackupStorageDestination = BackupStorageDestination.LOCAL
    ): Result<Unit> = withContext(Dispatchers.IO) {
        runCatching {
            val savedUri = getSavedBackupUri(destination)
                ?: throw IllegalStateException("Primero elige una ubicacion para el respaldo")
            val targetUri = resolveSavedTargetUri(savedUri)
            exportBackup(
                targetUri,
                persistPermission = false,
                updateSavedUri = false,
                destination = destination
            ).getOrElse { error ->
                if (shouldResetSavedLocation(error)) {
                    prefs.clearBackupUri(destination)
                    throw IllegalStateException("La ubicacion de respaldo ya no es valida. Configura la ubicacion nuevamente.")
                }
                throw error
            }
        }
    }

    suspend fun importBackup(sourceUri: Uri): Result<Unit> = withContext(Dispatchers.IO) {
        runCatching {
            val rawBytes = context.contentResolver.openInputStream(sourceUri)?.use { it.readBytes() }
                ?: error("No se pudo abrir archivo de backup")
            importBackupFromBytes(rawBytes).getOrThrow()
        }
    }

    suspend fun exportBackupBytes(): Result<ByteArray> = withContext(Dispatchers.IO) {
        runCatching {
            val plainBytes = DatabaseBackupManager.exportDatabaseBytes(context).getOrThrow()
            val password = requireBackupPassword()
            BackupCryptoManager.encryptPortable(plainBytes, password)
        }.recoverCatching { error ->
            throw mapCryptoError(error, onExport = true)
        }
    }

    suspend fun importBackupFromBytes(rawBytes: ByteArray): Result<Unit> = withContext(Dispatchers.IO) {
        runCatching {
            val password = prefs.backupPassword.first()
            val plainBytes = BackupCryptoManager.decryptAuto(rawBytes, password)
            DatabaseBackupManager.importDatabaseFromBytes(context, plainBytes).getOrThrow()
        }.recoverCatching { error ->
            throw mapCryptoError(error, onExport = false)
        }
    }

    fun observeHasBackupPassword() = prefs.backupPassword.map { !it.isNullOrBlank() }

    suspend fun hasBackupPassword(): Boolean = !prefs.backupPassword.first().isNullOrBlank()

    suspend fun setBackupPassword(password: String) {
        require(password.length >= 6) { "La clave de respaldo debe tener al menos 6 caracteres" }
        prefs.saveBackupPassword(password)
    }

    suspend fun clearBackupPassword() {
        prefs.clearBackupPassword()
    }

    suspend fun getSavedBackupUri(destination: BackupStorageDestination = BackupStorageDestination.LOCAL): Uri? {
        val value = prefs.backupUri(destination).first() ?: return null
        return runCatching { Uri.parse(value) }.getOrNull()
    }

    suspend fun hasSavedLocation(destination: BackupStorageDestination = BackupStorageDestination.LOCAL): Boolean {
        val uri = getSavedBackupUri(destination) ?: return false
        if (destination == BackupStorageDestination.DRIVE && !isDriveUri(uri)) return false
        return true
    }

    fun isDriveUri(uri: Uri): Boolean {
        val authority = uri.authority.orEmpty().lowercase()
        return authority.contains("com.google.android.apps.docs")
    }

    fun observeLastBackupTimestamp() = prefs.lastBackupTimestamp

    suspend fun markBackupSuccessNow() {
        prefs.saveLastBackupTimestamp(System.currentTimeMillis())
    }

    private fun persistUriPermissions(uri: Uri) {
        val resolver = context.contentResolver
        val candidates = listOf(
            IntentFlags.READ or IntentFlags.WRITE,
            IntentFlags.READ,
            IntentFlags.WRITE
        )
        candidates.forEach { flags ->
            runCatching {
                resolver.takePersistableUriPermission(uri, flags)
            }
        }
    }

    private fun shouldResetSavedLocation(error: Throwable): Boolean {
        return error is SecurityException ||
            error is FileNotFoundException ||
            (error is IllegalStateException && error.message?.contains("no valida", ignoreCase = true) == true)
    }

    private suspend fun requireBackupPassword(): String {
        return prefs.backupPassword.first()
            ?.takeIf { it.isNotBlank() }
            ?: throw IllegalStateException("Configura tu clave de respaldo para crear backups portables")
    }

    private fun mapCryptoError(error: Throwable, onExport: Boolean): Throwable {
        if (error is IllegalStateException) return error
        if (error is GeneralSecurityException) {
            return IllegalStateException(
                if (onExport) "No se pudo cifrar el respaldo" else "No se pudo descifrar el respaldo",
                error
            )
        }
        val message = error.message.orEmpty().lowercase()
        if ("tag mismatch" in message || "mac check" in message || "padding" in message) {
            return IllegalStateException("El respaldo cifrado es invalido o fue alterado", error)
        }
        return error
    }

    private fun createOrFindBackupFileInFolder(folderUri: Uri): Uri {
        val folder = DocumentFile.fromTreeUri(context, folderUri)
            ?: throw IllegalStateException("Error al crear respaldo")
        require(folder.isDirectory) { "Error al crear respaldo" }
        val existing = folder.findFile(BACKUP_FILE_NAME)
        val file = existing ?: folder.createFile("application/octet-stream", BACKUP_FILE_NAME)
        return file?.uri ?: throw IllegalStateException("Error al crear respaldo")
    }

    private fun resolveSavedTargetUri(savedUri: Uri): Uri {
        return if (DocumentsContract.isTreeUri(savedUri)) {
            createOrFindBackupFileInFolder(savedUri)
        } else {
            savedUri
        }
    }

    object IntentFlags {
        const val READ: Int = Intent.FLAG_GRANT_READ_URI_PERMISSION
        const val WRITE: Int = Intent.FLAG_GRANT_WRITE_URI_PERMISSION
    }

    companion object {
        const val BACKUP_FILE_NAME = "prestamos_backup.db"

        fun restartApplication(context: Context) {
            val launchIntent = context.packageManager.getLaunchIntentForPackage(context.packageName)
                ?.addFlags(Intent.FLAG_ACTIVITY_NEW_TASK or Intent.FLAG_ACTIVITY_CLEAR_TASK)
                ?: return
            context.startActivity(launchIntent)
        }
    }
}
