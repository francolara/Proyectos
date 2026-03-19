package com.prestamos.app.data.backup

import android.content.Context
import android.content.Intent
import android.net.Uri
import android.provider.DocumentsContract
import androidx.documentfile.provider.DocumentFile
import androidx.room.withTransaction
import com.prestamos.app.data.local.AppDatabase
import com.prestamos.app.data.local.entity.ClienteEntity
import com.prestamos.app.data.local.entity.CuotaEntity
import com.prestamos.app.data.local.entity.EstadoCuota
import com.prestamos.app.data.local.entity.EstadoPrestamo
import com.prestamos.app.data.local.entity.Moneda
import com.prestamos.app.data.local.entity.PagoEntity
import com.prestamos.app.data.local.entity.PrestamoEntity
import com.prestamos.app.data.local.entity.TipoPago
import java.io.FileNotFoundException
import kotlinx.coroutines.Dispatchers
import kotlinx.coroutines.flow.first
import kotlinx.coroutines.withContext
import org.json.JSONArray
import org.json.JSONObject

class BackupManager(private val context: Context) {
    private val db = AppDatabase.getInstance(context)
    private val prefs = BackupPreferences(context)

    suspend fun exportBackup(targetUri: Uri, persistPermission: Boolean): Result<Unit> = withContext(Dispatchers.IO) {
        runCatching {
            if (persistPermission) persistUriPermissions(targetUri)
            val data = buildBackupData()
            val json = serialize(data)
            val stream = context.contentResolver.openOutputStream(targetUri, "wt")
                ?: context.contentResolver.openOutputStream(targetUri, "w")
            stream?.use { output ->
                output.write(json.toByteArray(Charsets.UTF_8))
            } ?: error("Error al crear respaldo")
            prefs.saveBackupUri(targetUri.toString())
            prefs.saveLastBackupTimestamp(data.fechaBackup)
        }
    }

    suspend fun exportBackupToFolder(folderUri: Uri, persistPermission: Boolean): Result<Unit> = withContext(Dispatchers.IO) {
        runCatching {
            if (persistPermission) persistUriPermissions(folderUri)
            val fileUri = createOrFindBackupFileInFolder(folderUri)
            exportBackup(fileUri, persistPermission = false).getOrThrow()
            prefs.saveBackupUri(folderUri.toString())
        }
    }
    suspend fun exportBackupToSavedLocation(): Result<Unit> = withContext(Dispatchers.IO) {
        runCatching {
            val savedUri = getSavedBackupUri()
                ?: throw IllegalStateException("Primero elige una ubicacion para el respaldo")
            val targetUri = resolveSavedTargetUri(savedUri)
            exportBackup(targetUri, persistPermission = false).getOrElse { error ->
                if (shouldResetSavedLocation(error)) {
                    prefs.clearBackupUri()
                    throw IllegalStateException("La ubicacion de respaldo ya no es valida. Configura la ubicacion nuevamente.")
                }
                throw error
            }
        }
    }

    suspend fun importBackup(sourceUri: Uri): Result<Unit> = withContext(Dispatchers.IO) {
        runCatching {
            val rawJson = context.contentResolver.openInputStream(sourceUri)?.bufferedReader(Charsets.UTF_8)?.use { it.readText() }
                ?: error("Archivo invalido")
            val parsed = deserialize(rawJson)

            db.withTransaction {
                db.pagoDao().eliminarTodos()
                db.cuotaDao().eliminarTodos()
                db.prestamoDao().eliminarTodos()
                db.clienteDao().eliminarTodos()

                if (parsed.clientes.isNotEmpty()) db.clienteDao().insertarTodos(parsed.clientes)
                if (parsed.prestamos.isNotEmpty()) db.prestamoDao().insertarTodos(parsed.prestamos)
                if (parsed.cuotas.isNotEmpty()) db.cuotaDao().insertarTodos(parsed.cuotas)
                if (parsed.pagos.isNotEmpty()) db.pagoDao().insertarTodos(parsed.pagos)
            }
        }
    }

    suspend fun getSavedBackupUri(): Uri? {
        val value = prefs.backupUri.first() ?: return null
        return runCatching { Uri.parse(value) }.getOrNull()
    }

    suspend fun hasSavedLocation(): Boolean = getSavedBackupUri() != null

    fun observeLastBackupTimestamp() = prefs.lastBackupTimestamp

    private suspend fun buildBackupData(): BackupData {
        return BackupData(
            version = BACKUP_VERSION,
            fechaBackup = System.currentTimeMillis(),
            clientes = db.clienteDao().listarTodosInterno(),
            prestamos = db.prestamoDao().listarTodosInterno(),
            cuotas = db.cuotaDao().listarTodasInterno(),
            pagos = db.pagoDao().listarTodosInterno()
        )
    }

    private fun serialize(data: BackupData): String {
        val root = JSONObject()
            .put("version", data.version)
            .put("fechaBackup", data.fechaBackup)
            .put("clientes", JSONArray().apply { data.clientes.forEach { put(clienteToJson(it)) } })
            .put("prestamos", JSONArray().apply { data.prestamos.forEach { put(prestamoToJson(it)) } })
            .put("cuotas", JSONArray().apply { data.cuotas.forEach { put(cuotaToJson(it)) } })
            .put("pagos", JSONArray().apply { data.pagos.forEach { put(pagoToJson(it)) } })
        return root.toString(2)
    }

    private fun deserialize(rawJson: String): BackupData {
        val root = runCatching { JSONObject(rawJson) }.getOrElse { throw IllegalArgumentException("Archivo invalido") }
        val version = root.optInt("version", -1)
        require(version == BACKUP_VERSION) { "Archivo invalido" }

        val fechaBackup = root.optLong("fechaBackup", -1L)
        require(fechaBackup > 0L) { "Archivo invalido" }

        val clientesJson = root.optJSONArray("clientes") ?: throw IllegalArgumentException("Archivo invalido")
        val prestamosJson = root.optJSONArray("prestamos") ?: throw IllegalArgumentException("Archivo invalido")
        val cuotasJson = root.optJSONArray("cuotas") ?: throw IllegalArgumentException("Archivo invalido")
        val pagosJson = root.optJSONArray("pagos") ?: throw IllegalArgumentException("Archivo invalido")

        val clientes = (0 until clientesJson.length()).map { idx -> jsonToCliente(clientesJson.getJSONObject(idx)) }
        val prestamos = (0 until prestamosJson.length()).map { idx -> jsonToPrestamo(prestamosJson.getJSONObject(idx)) }
        val cuotas = (0 until cuotasJson.length()).map { idx -> jsonToCuota(cuotasJson.getJSONObject(idx)) }
        val pagos = (0 until pagosJson.length()).map { idx -> jsonToPago(pagosJson.getJSONObject(idx)) }

        return BackupData(
            version = version,
            fechaBackup = fechaBackup,
            clientes = clientes,
            prestamos = prestamos,
            cuotas = cuotas,
            pagos = pagos
        )
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

    private fun createOrFindBackupFileInFolder(folderUri: Uri): Uri {
        val folder = DocumentFile.fromTreeUri(context, folderUri)
            ?: throw IllegalStateException("Error al crear respaldo")
        require(folder.isDirectory) { "Error al crear respaldo" }
        val existing = folder.findFile(BACKUP_FILE_NAME)
        val file = existing ?: folder.createFile("application/json", BACKUP_FILE_NAME)
        return file?.uri ?: throw IllegalStateException("Error al crear respaldo")
    }

    private fun resolveSavedTargetUri(savedUri: Uri): Uri {
        return if (DocumentsContract.isTreeUri(savedUri)) {
            createOrFindBackupFileInFolder(savedUri)
        } else {
            savedUri
        }
    }

    private fun clienteToJson(item: ClienteEntity): JSONObject = JSONObject()
        .put("idCliente", item.idCliente)
        .put("nombre", item.nombre)
        .put("apellido", item.apellido)
        .put("documentoIdentidad", item.documentoIdentidad)
        .put("direccion", item.direccion)
        .put("telefono", item.telefono)
        .put("fechaRegistro", item.fechaRegistro)
        .put("fechaModificacion", item.fechaModificacion)

    private fun prestamoToJson(item: PrestamoEntity): JSONObject = JSONObject()
        .put("idPrestamo", item.idPrestamo)
        .put("idCliente", item.idCliente)
        .put("montoPrestado", item.montoPrestado)
        .put("interes", item.interes)
        .put("moneda", item.moneda.name)
        .put("tipoPago", item.tipoPago.name)
        .put("cantidadCuotas", item.cantidadCuotas)
        .put("fechaPrimeraCuota", item.fechaPrimeraCuota)
        .put("montoTotalPrestamo", item.montoTotalPrestamo)
        .put("montoCuota", item.montoCuota)
        .put("estadoPrestamo", item.estadoPrestamo.name)
        .put("fechaRegistro", item.fechaRegistro)
        .put("fechaModificacion", item.fechaModificacion)

    private fun cuotaToJson(item: CuotaEntity): JSONObject = JSONObject()
        .put("idCuota", item.idCuota)
        .put("idPrestamo", item.idPrestamo)
        .put("numeroCuota", item.numeroCuota)
        .put("fechaVencimiento", item.fechaVencimiento)
        .put("montoCuota", item.montoCuota)
        .put("montoPagado", item.montoPagado)
        .put("saldoPendiente", item.saldoPendiente)
        .put("estadoCuota", item.estadoCuota.name)
        .put("fechaRegistro", item.fechaRegistro)
        .put("fechaModificacion", item.fechaModificacion)

    private fun pagoToJson(item: PagoEntity): JSONObject = JSONObject()
        .put("idPago", item.idPago)
        .put("idPrestamo", item.idPrestamo)
        .put("idCuota", item.idCuota)
        .put("fechaPago", item.fechaPago)
        .put("montoAbono", item.montoAbono)
        .put("observacion", item.observacion)
        .put("fechaRegistro", item.fechaRegistro)
        .put("fechaModificacion", item.fechaModificacion)

    private fun jsonToCliente(json: JSONObject): ClienteEntity = ClienteEntity(
        idCliente = json.optLongStrict("idCliente"),
        nombre = json.optStringStrict("nombre"),
        apellido = json.optStringStrict("apellido"),
        documentoIdentidad = json.optStringStrict("documentoIdentidad"),
        direccion = json.optStringStrict("direccion"),
        telefono = json.optStringStrict("telefono"),
        fechaRegistro = json.optLongStrict("fechaRegistro"),
        fechaModificacion = json.optLongStrict("fechaModificacion")
    )

    private fun jsonToPrestamo(json: JSONObject): PrestamoEntity = PrestamoEntity(
        idPrestamo = json.optLongStrict("idPrestamo"),
        idCliente = json.optLongStrict("idCliente"),
        montoPrestado = json.optDoubleStrict("montoPrestado"),
        interes = json.optDoubleStrict("interes"),
        moneda = enumValueOfOrThrow<Moneda>(json.optStringStrict("moneda")),
        tipoPago = enumValueOfOrThrow<TipoPago>(json.optStringStrict("tipoPago")),
        cantidadCuotas = json.optIntStrict("cantidadCuotas"),
        fechaPrimeraCuota = json.optLongStrict("fechaPrimeraCuota"),
        montoTotalPrestamo = json.optDoubleStrict("montoTotalPrestamo"),
        montoCuota = json.optDoubleStrict("montoCuota"),
        estadoPrestamo = enumValueOfOrThrow<EstadoPrestamo>(json.optStringStrict("estadoPrestamo")),
        fechaRegistro = json.optLongStrict("fechaRegistro"),
        fechaModificacion = json.optLongStrict("fechaModificacion")
    )

    private fun jsonToCuota(json: JSONObject): CuotaEntity = CuotaEntity(
        idCuota = json.optLongStrict("idCuota"),
        idPrestamo = json.optLongStrict("idPrestamo"),
        numeroCuota = json.optIntStrict("numeroCuota"),
        fechaVencimiento = json.optLongStrict("fechaVencimiento"),
        montoCuota = json.optDoubleStrict("montoCuota"),
        montoPagado = json.optDoubleStrict("montoPagado"),
        saldoPendiente = json.optDoubleStrict("saldoPendiente"),
        estadoCuota = enumValueOfOrThrow<EstadoCuota>(json.optStringStrict("estadoCuota")),
        fechaRegistro = json.optLongStrict("fechaRegistro"),
        fechaModificacion = json.optLongStrict("fechaModificacion")
    )

    private fun jsonToPago(json: JSONObject): PagoEntity = PagoEntity(
        idPago = json.optLongStrict("idPago"),
        idPrestamo = json.optLongStrict("idPrestamo"),
        idCuota = json.optLongStrict("idCuota"),
        fechaPago = json.optLongStrict("fechaPago"),
        montoAbono = json.optDoubleStrict("montoAbono"),
        observacion = if (json.isNull("observacion")) null else json.optString("observacion"),
        fechaRegistro = json.optLongStrict("fechaRegistro"),
        fechaModificacion = json.optLongStrict("fechaModificacion")
    )

    private inline fun <reified T : Enum<T>> enumValueOfOrThrow(value: String): T {
        return runCatching { enumValueOf<T>(value) }.getOrElse { throw IllegalArgumentException("Archivo invalido") }
    }

    private fun JSONObject.optStringStrict(key: String): String {
        val value = optString(key, "")
        require(value.isNotBlank()) { "Archivo invalido" }
        return value
    }

    private fun JSONObject.optLongStrict(key: String): Long {
        require(has(key)) { "Archivo invalido" }
        return optLong(key, Long.MIN_VALUE).also { require(it != Long.MIN_VALUE) { "Archivo invalido" } }
    }

    private fun JSONObject.optIntStrict(key: String): Int {
        require(has(key)) { "Archivo invalido" }
        return optInt(key, Int.MIN_VALUE).also { require(it != Int.MIN_VALUE) { "Archivo invalido" } }
    }

    private fun JSONObject.optDoubleStrict(key: String): Double {
        require(has(key)) { "Archivo invalido" }
        return optDouble(key, Double.NaN).also { require(!it.isNaN()) { "Archivo invalido" } }
    }

    object IntentFlags {
        const val READ: Int = Intent.FLAG_GRANT_READ_URI_PERMISSION
        const val WRITE: Int = Intent.FLAG_GRANT_WRITE_URI_PERMISSION
    }

    companion object {
        const val BACKUP_FILE_NAME = "prestamos_backup.json"
        const val BACKUP_VERSION = 1

        fun restartApplication(context: Context) {
            val launchIntent = context.packageManager.getLaunchIntentForPackage(context.packageName)
                ?.addFlags(Intent.FLAG_ACTIVITY_NEW_TASK or Intent.FLAG_ACTIVITY_CLEAR_TASK)
                ?: return
            context.startActivity(launchIntent)
        }
    }
}
