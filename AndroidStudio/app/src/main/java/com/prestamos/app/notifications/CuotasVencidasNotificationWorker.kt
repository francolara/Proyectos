package com.prestamos.app.notifications

import android.Manifest
import android.app.NotificationChannel
import android.app.NotificationManager
import android.app.PendingIntent
import android.content.Context
import android.content.Intent
import android.content.pm.PackageManager
import android.os.Build
import androidx.core.app.NotificationCompat
import androidx.core.app.NotificationManagerCompat
import androidx.core.content.ContextCompat
import androidx.work.CoroutineWorker
import androidx.work.WorkerParameters
import com.prestamos.app.MainActivity
import com.prestamos.app.R
import com.prestamos.app.data.local.AppDatabase
import java.time.LocalDate
import java.time.ZoneId

class CuotasVencidasNotificationWorker(
    appContext: Context,
    workerParams: WorkerParameters
) : CoroutineWorker(appContext, workerParams) {

    override suspend fun doWork(): Result {
        ensureChannel()

        val endToday = LocalDate.now()
            .plusDays(1)
            .atStartOfDay(ZoneId.systemDefault())
            .toInstant()
            .toEpochMilli() - 1

        val db = AppDatabase.getInstance(applicationContext)
        val cuotasVencidas = db.cuotaDao().listarVencidasParaNotificacion(endToday)
        if (cuotasVencidas.isEmpty()) {
            return Result.success()
        }

        if (Build.VERSION.SDK_INT >= Build.VERSION_CODES.TIRAMISU &&
            ContextCompat.checkSelfPermission(applicationContext, Manifest.permission.POST_NOTIFICATIONS) != PackageManager.PERMISSION_GRANTED
        ) {
            return Result.success()
        }

        val prestamoDao = db.prestamoDao()
        val clienteDao = db.clienteDao()

        val preview = cuotasVencidas.take(3).joinToString("\n") { cuota ->
            val prestamo = prestamoDao.obtenerPorId(cuota.idPrestamo)
            val cliente = prestamo?.let { clienteDao.obtenerPorId(it.idCliente) }
            val nombreCliente = "${cliente?.nombre.orEmpty()} ${cliente?.apellido.orEmpty()}".trim().ifBlank { "Cliente" }
            "• $nombreCliente - Préstamo #${cuota.idPrestamo} cuota ${cuota.numeroCuota}"
        }

        val intent = Intent(applicationContext, MainActivity::class.java)
        val flags = PendingIntent.FLAG_UPDATE_CURRENT or PendingIntent.FLAG_IMMUTABLE
        val pendingIntent = PendingIntent.getActivity(applicationContext, 2001, intent, flags)

        val contentText = if (preview.isBlank()) {
            "Tienes ${cuotasVencidas.size} cuota(s) vencida(s) o por vencer hoy sin pagar."
        } else {
            "$preview\n\nTotal pendientes: ${cuotasVencidas.size}"
        }

        val notification = NotificationCompat.Builder(applicationContext, CHANNEL_ID)
            .setSmallIcon(R.mipmap.ic_launcher)
            .setContentTitle("Recordatorio de cuotas vencidas")
            .setContentText("Tienes ${cuotasVencidas.size} cuota(s) vencida(s).")
            .setStyle(NotificationCompat.BigTextStyle().bigText(contentText))
            .setPriority(NotificationCompat.PRIORITY_HIGH)
            .setAutoCancel(true)
            .setContentIntent(pendingIntent)
            .build()

        NotificationManagerCompat.from(applicationContext).notify(NOTIFICATION_ID, notification)
        return Result.success()
    }

    private fun ensureChannel() {
        if (Build.VERSION.SDK_INT < Build.VERSION_CODES.O) return
        val manager = applicationContext.getSystemService(Context.NOTIFICATION_SERVICE) as NotificationManager
        val channel = NotificationChannel(
            CHANNEL_ID,
            "Cuotas vencidas",
            NotificationManager.IMPORTANCE_HIGH
        ).apply {
            description = "Notificaciones diarias de cuotas vencidas"
        }
        manager.createNotificationChannel(channel)
    }

    companion object {
        private const val CHANNEL_ID = "cuotas_vencidas_channel"
        private const val NOTIFICATION_ID = 2027
    }
}
