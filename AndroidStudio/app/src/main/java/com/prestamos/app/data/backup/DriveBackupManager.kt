package com.prestamos.app.data.backup

import android.content.Context
import com.google.api.client.googleapis.extensions.android.gms.auth.GoogleAccountCredential
import com.google.api.client.http.ByteArrayContent
import com.google.api.client.http.javanet.NetHttpTransport
import com.google.api.client.json.gson.GsonFactory
import com.google.api.services.drive.Drive
import com.google.api.services.drive.DriveScopes
import com.google.api.services.drive.model.File
import com.google.android.gms.auth.api.signin.GoogleSignInAccount
import java.io.ByteArrayOutputStream
import kotlinx.coroutines.Dispatchers
import kotlinx.coroutines.withContext

class DriveBackupManager(
    private val context: Context,
    private val backupManager: BackupManager
) {
    private val backupFolderName = "AppPrestamos"

    suspend fun subirRespaldo(account: GoogleSignInAccount): Result<Unit> = withContext(Dispatchers.IO) {
        runCatching {
            val service = createDriveService(account)
            val folderId = obtenerOCrearCarpeta(service)
            val existingFiles = buscarArchivosRespaldo(service, folderId)
            val existing = existingFiles.firstOrNull()
            val backupBytes = backupManager.exportBackupBytes().getOrThrow()
            val mediaContent = ByteArrayContent("application/octet-stream", backupBytes)

            if (existing != null) {
                service.files()
                    .update(existing.id, null, mediaContent)
                    .setFields("id, modifiedTime")
                    .execute()
            } else {
                val metadata = File().apply {
                    name = BackupManager.BACKUP_FILE_NAME
                    parents = listOf(folderId)
                    mimeType = "application/octet-stream"
                }
                service.files()
                    .create(metadata, mediaContent)
                    .setFields("id")
                    .execute()
            }
            if (existingFiles.size > 1) {
                existingFiles.drop(1).forEach { duplicate ->
                    runCatching { service.files().delete(duplicate.id).execute() }
                }
            }
            backupManager.markBackupSuccessNow()
            Unit
        }
    }

    suspend fun restaurarRespaldo(account: GoogleSignInAccount): Result<Unit> = withContext(Dispatchers.IO) {
        runCatching {
            val service = createDriveService(account)
            val folderId = obtenerOCrearCarpeta(service)
            val existing = buscarArchivosRespaldo(service, folderId).firstOrNull()
                ?: throw IllegalStateException("No hay respaldo en Drive")

            val output = ByteArrayOutputStream()
            service.files().get(existing.id).executeMediaAndDownloadTo(output)
            val rawBytes = output.toByteArray()
            backupManager.importBackupFromBytes(rawBytes).getOrThrow()
        }
    }

    private fun obtenerOCrearCarpeta(service: Drive): String {
        val escapedFolderName = backupFolderName.replace("'", "\\'")
        val query = "mimeType='application/vnd.google-apps.folder' and name='$escapedFolderName' and 'root' in parents and trashed=false"
        val result = service.files()
            .list()
            .setSpaces("drive")
            .setQ(query)
            .setFields("files(id,name,modifiedTime)")
            .setOrderBy("modifiedTime desc")
            .setPageSize(1)
            .execute()
        val existingFolder = result.files?.firstOrNull()
        if (existingFolder != null) return existingFolder.id

        val folderMetadata = File().apply {
            name = backupFolderName
            mimeType = "application/vnd.google-apps.folder"
            parents = listOf("root")
        }
        return service.files()
            .create(folderMetadata)
            .setFields("id")
            .execute()
            .id
    }

    private fun buscarArchivosRespaldo(service: Drive, folderId: String): List<File> {
        val escapedFileName = BackupManager.BACKUP_FILE_NAME.replace("'", "\\'")
        val query = "name='$escapedFileName' and '$folderId' in parents and trashed=false"
        val result = service.files()
            .list()
            .setSpaces("drive")
            .setQ(query)
            .setFields("files(id,name,modifiedTime)")
            .setOrderBy("modifiedTime desc")
            .setPageSize(100)
            .execute()
        return result.files ?: emptyList()
    }

    private fun createDriveService(account: GoogleSignInAccount): Drive {
        val credential = GoogleAccountCredential.usingOAuth2(
            context,
            setOf(DriveScopes.DRIVE_FILE)
        ).apply {
            selectedAccount = account.account
        }

        return Drive.Builder(
            NetHttpTransport(),
            GsonFactory.getDefaultInstance(),
            credential
        ).setApplicationName("AppPrestamos").build()
    }
}
