package com.prestamos.app.data.local

import android.content.Context
import android.content.Intent
import android.database.sqlite.SQLiteDatabase
import android.net.Uri
import kotlinx.coroutines.Dispatchers
import kotlinx.coroutines.withContext
import java.io.File
import java.io.FileInputStream
import java.io.FileOutputStream
import java.time.LocalDateTime
import java.time.format.DateTimeFormatter
import kotlin.system.exitProcess

object DatabaseBackupManager {
    private val backupDateFormat = DateTimeFormatter.ofPattern("yyyy_MM_dd_HHmm")

    fun buildBackupFileName(now: LocalDateTime = LocalDateTime.now()): String {
        return "backup_prestamos_${now.format(backupDateFormat)}.db"
    }

    suspend fun exportDatabase(context: Context, targetUri: Uri): Result<Unit> = withContext(Dispatchers.IO) {
        runCatching {
            val dbFile = context.getDatabasePath(AppDatabase.DATABASE_NAME)
            require(dbFile.exists()) { "No se encontró la base de datos local" }

            val snapshotFile = createConsistentSnapshot(context)
            try {
                context.contentResolver.openOutputStream(targetUri)?.use { output ->
                    FileInputStream(snapshotFile).use { input ->
                        input.copyTo(output)
                    }
                } ?: error("No se pudo abrir destino de exportación")
            } finally {
                snapshotFile.delete()
            }
            Unit
        }
    }

    suspend fun importDatabase(context: Context, sourceUri: Uri): Result<Unit> = withContext(Dispatchers.IO) {
        runCatching {
            val tempFile = File(context.cacheDir, "restore_temp.db")
            context.contentResolver.openInputStream(sourceUri)?.use { input ->
                FileOutputStream(tempFile).use { output ->
                    input.copyTo(output)
                }
            } ?: error("No se pudo abrir archivo de backup")

            validarBackup(tempFile)

            val dbFile = context.getDatabasePath(AppDatabase.DATABASE_NAME)
            val dbDir = dbFile.parentFile
            if (dbDir != null && !dbDir.exists()) dbDir.mkdirs()

            AppDatabase.closeInstance()

            File(dbFile.absolutePath + "-wal").delete()
            File(dbFile.absolutePath + "-shm").delete()

            FileInputStream(tempFile).use { input ->
                FileOutputStream(dbFile, false).use { output ->
                    input.copyTo(output)
                }
            }

            tempFile.delete()
            Unit
        }
    }

    fun restartApplication(context: Context) {
        val launchIntent = context.packageManager.getLaunchIntentForPackage(context.packageName)
            ?.addFlags(Intent.FLAG_ACTIVITY_NEW_TASK or Intent.FLAG_ACTIVITY_CLEAR_TASK)
            ?: return
        context.startActivity(launchIntent)
        exitProcess(0)
    }

    private fun createConsistentSnapshot(context: Context): File {
        val snapshotFile = File(context.cacheDir, "snapshot_${AppDatabase.DATABASE_NAME}")
        if (snapshotFile.exists()) snapshotFile.delete()

        val db = AppDatabase.getInstance(context).openHelper.writableDatabase
        db.query("PRAGMA wal_checkpoint(TRUNCATE)").close()

        val snapshotPath = snapshotFile.absolutePath.replace("'", "''")
        db.execSQL("VACUUM INTO '$snapshotPath'")

        require(snapshotFile.exists() && snapshotFile.length() > 0L) {
            "No se pudo generar snapshot de backup"
        }
        return snapshotFile
    }

    private fun validarBackup(file: File) {
        require(file.exists() && file.length() > 0L) { "El archivo de backup está vacío o no existe" }
        val db = SQLiteDatabase.openDatabase(file.absolutePath, null, SQLiteDatabase.OPEN_READONLY)
        db.rawQuery("PRAGMA integrity_check", null).use { cursor ->
            require(cursor.moveToFirst() && cursor.getString(0).equals("ok", ignoreCase = true)) {
                "El backup está corrupto"
            }
        }

        db.rawQuery("PRAGMA user_version", null).use { cursor ->
            require(cursor.moveToFirst()) { "No se pudo validar versión del backup" }
            val backupVersion = cursor.getInt(0)
            require(backupVersion == AppDatabase.DATABASE_VERSION) {
                "El backup corresponde a versión $backupVersion y la app requiere ${AppDatabase.DATABASE_VERSION}"
            }
        }
        db.close()
    }
}
