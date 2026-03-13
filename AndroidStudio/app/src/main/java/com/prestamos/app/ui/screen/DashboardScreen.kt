package com.prestamos.app.ui.screen

import android.content.Intent
import android.widget.Toast
import androidx.compose.foundation.background
import androidx.compose.foundation.clickable
import androidx.compose.foundation.layout.Arrangement
import androidx.compose.foundation.layout.Column
import androidx.compose.foundation.layout.Row
import androidx.compose.foundation.layout.Spacer
import androidx.compose.foundation.layout.fillMaxSize
import androidx.compose.foundation.layout.fillMaxWidth
import androidx.compose.foundation.layout.height
import androidx.compose.foundation.layout.padding
import androidx.compose.foundation.lazy.LazyColumn
import androidx.compose.foundation.lazy.items
import androidx.compose.foundation.shape.RoundedCornerShape
import androidx.compose.material3.Card
import androidx.compose.material3.MaterialTheme
import androidx.compose.material3.Switch
import androidx.compose.material3.Text
import androidx.compose.runtime.Composable
import androidx.compose.runtime.getValue
import androidx.compose.runtime.mutableStateOf
import androidx.compose.runtime.remember
import androidx.compose.runtime.setValue
import androidx.compose.ui.platform.LocalContext
import androidx.compose.ui.Alignment
import androidx.compose.ui.Modifier
import androidx.compose.ui.graphics.Color
import androidx.compose.ui.unit.dp
import androidx.core.content.FileProvider
import androidx.lifecycle.compose.collectAsStateWithLifecycle
import com.prestamos.app.ui.screen.export.createDashboardDetalleImage
import com.prestamos.app.ui.screen.export.createDashboardDetallePdf
import com.prestamos.app.ui.viewmodel.DashboardViewModel
import com.prestamos.app.util.toDateString
import com.prestamos.app.util.toMoney
import java.io.File

private enum class DashboardDetalle {
    CAPITAL,
    PENDIENTE,
    COBRADO_HOY,
    VENCIDAS,
    CUOTAS_PAGADAS,
    CUOTAS_PENDIENTES,
    CUOTAS_PARCIALES,
    CUOTAS_VENCIDAS
}

@Composable
fun DashboardScreen(
    viewModel: DashboardViewModel,
    isDarkMode: Boolean,
    onToggleDarkMode: (Boolean) -> Unit
) {
    val context = LocalContext.current
    val state by viewModel.uiState.collectAsStateWithLifecycle()
    var detalleSeleccionado by remember { mutableStateOf<DashboardDetalle?>(null) }

    LazyColumn(
        modifier = Modifier
            .fillMaxSize()
            .padding(16.dp),
        verticalArrangement = Arrangement.spacedBy(12.dp)
    ) {
        item {
            Row(
                modifier = Modifier.fillMaxWidth(),
                horizontalArrangement = Arrangement.SpaceBetween,
                verticalAlignment = Alignment.CenterVertically
            ) {
                Column {
                    Text("Resumen general", style = MaterialTheme.typography.headlineSmall)
                    Text("Fecha: ${System.currentTimeMillis().toDateString()}")
                }
                Row(verticalAlignment = Alignment.CenterVertically, horizontalArrangement = Arrangement.spacedBy(8.dp)) {
                    Text("Modo oscuro")
                    Switch(checked = isDarkMode, onCheckedChange = onToggleDarkMode)
                }
            }
        }

        item {
            Row(horizontalArrangement = Arrangement.spacedBy(8.dp), modifier = Modifier.fillMaxWidth()) {
                DashboardCard("Capital prestado", state.capitalPrestado.toMoney(state.monedaReferencial), Modifier.weight(1f)) {
                    detalleSeleccionado = DashboardDetalle.CAPITAL
                }
                DashboardCard("Saldo pendiente", state.saldoPendiente.toMoney(state.monedaReferencial), Modifier.weight(1f)) {
                    detalleSeleccionado = DashboardDetalle.PENDIENTE
                }
            }
            Spacer(modifier = Modifier.height(8.dp))
            Row(horizontalArrangement = Arrangement.spacedBy(8.dp), modifier = Modifier.fillMaxWidth()) {
                DashboardCard("Cobrado hoy", state.cobradoHoy.toMoney(state.monedaReferencial), Modifier.weight(1f)) {
                    detalleSeleccionado = DashboardDetalle.COBRADO_HOY
                }
                DashboardCard("Cuotas vencidas", state.cuotasVencidas.toString(), Modifier.weight(1f)) {
                    detalleSeleccionado = DashboardDetalle.VENCIDAS
                }
            }
        }

        item {
            Text("Gráfico comparativo", style = MaterialTheme.typography.titleMedium)
            val total = (state.capitalPrestado + state.cobradoHoy + state.saldoPendiente).coerceAtLeast(1.0)
            Row(modifier = Modifier.fillMaxWidth(), horizontalArrangement = Arrangement.spacedBy(8.dp)) {
                BarSegment("Prestado", (state.capitalPrestado / total).toFloat(), Color(0xFF1565C0)) { detalleSeleccionado = DashboardDetalle.CAPITAL }
                BarSegment("Cobrado", (state.cobradoHoy / total).toFloat(), Color(0xFF2E7D32)) { detalleSeleccionado = DashboardDetalle.COBRADO_HOY }
                BarSegment("Pendiente", (state.saldoPendiente / total).toFloat(), Color(0xFFF57C00)) { detalleSeleccionado = DashboardDetalle.PENDIENTE }
            }
        }

        item {
            Text("Estado de cuotas", style = MaterialTheme.typography.titleMedium)
            Row(modifier = Modifier.fillMaxWidth(), horizontalArrangement = Arrangement.spacedBy(8.dp)) {
                DashboardCard("Pagadas", (state.estadoCuotas["Pagadas"] ?: 0).toString(), Modifier.weight(1f)) { detalleSeleccionado = DashboardDetalle.CUOTAS_PAGADAS }
                DashboardCard("Pendientes", (state.estadoCuotas["Pendientes"] ?: 0).toString(), Modifier.weight(1f)) { detalleSeleccionado = DashboardDetalle.CUOTAS_PENDIENTES }
            }
            Spacer(Modifier.height(8.dp))
            Row(modifier = Modifier.fillMaxWidth(), horizontalArrangement = Arrangement.spacedBy(8.dp)) {
                DashboardCard("Parciales", (state.estadoCuotas["Parciales"] ?: 0).toString(), Modifier.weight(1f)) { detalleSeleccionado = DashboardDetalle.CUOTAS_PARCIALES }
                DashboardCard("Vencidas", (state.estadoCuotas["Vencidas"] ?: 0).toString(), Modifier.weight(1f)) { detalleSeleccionado = DashboardDetalle.CUOTAS_VENCIDAS }
            }
        }

        item {
            Text("Próximos vencimientos", style = MaterialTheme.typography.titleMedium)
            if (state.proximosVencimientos.isEmpty()) {
                Text("No hay próximos vencimientos")
            }
        }
        items(state.proximosVencimientos) { cuota ->
            Card(modifier = Modifier.fillMaxWidth()) {
                Column(Modifier.padding(10.dp)) {
                    Text(cuota.cliente)
                    Text("Préstamo #${cuota.idPrestamo} | Cuota ${cuota.numeroCuota}")
                    Text("Vence ${cuota.fechaVencimiento.toDateString()}")
                    Text("Saldo: ${cuota.saldoPendiente.toMoney(cuota.moneda)}")
                }
            }
        }

        item {
            Text("Últimos pagos", style = MaterialTheme.typography.titleMedium)
            if (state.ultimosPagos.isEmpty()) {
                Text("No hay pagos registrados")
            }
        }
        items(state.ultimosPagos) { pago ->
            Card(modifier = Modifier.fillMaxWidth()) {
                Column(Modifier.padding(10.dp)) {
                    Text(pago.cliente)
                    Text("Préstamo #${pago.idPrestamo} | Cuota ${pago.numeroCuota}")
                    Text("Fecha: ${pago.fechaPago.toDateString()}")
                    Text("Abono: ${pago.montoAbono.toMoney(pago.moneda)}")
                }
            }
        }
    }

    detalleSeleccionado?.let {
        val detalle = it.toDetalleInfo(state)
        DetalleDashboardDialog(
            detalle = detalle,
            onClose = { detalleSeleccionado = null },
            onShareText = {
                compartirTextoDetalle(context, detalle)
            },
            onShareImage = {
                runCatching {
                    createDashboardDetalleImage(context, detalle.title, detalle.message)
                }.onSuccess { file ->
                    compartirArchivoDetalle(context, file, "image/png")
                }.onFailure {
                    Toast.makeText(context, "No se pudo exportar imagen", Toast.LENGTH_SHORT).show()
                }
            },
            onSharePdf = {
                runCatching {
                    createDashboardDetallePdf(context, detalle.title, detalle.message)
                }.onSuccess { file ->
                    compartirArchivoDetalle(context, file, "application/pdf")
                }.onFailure {
                    Toast.makeText(context, "No se pudo exportar PDF", Toast.LENGTH_SHORT).show()
                }
            }
        )
    }
}

@Composable
private fun DashboardCard(title: String, value: String, modifier: Modifier = Modifier, onClick: () -> Unit) {
    Card(modifier = modifier.clickable(onClick = onClick)) {
        Column(Modifier.padding(12.dp)) {
            Text(title, style = MaterialTheme.typography.labelLarge)
            Text(value, style = MaterialTheme.typography.titleLarge)
        }
    }
}

@Composable
private fun BarSegment(label: String, ratio: Float, color: Color, onClick: () -> Unit) {
    val height = (80 * ratio.coerceIn(0.15f, 1f)).dp
    Column(horizontalAlignment = Alignment.CenterHorizontally, modifier = Modifier.clickable(onClick = onClick)) {
        Spacer(
            Modifier
                .height(height)
                .fillMaxWidth(0.28f)
                .background(color, RoundedCornerShape(8.dp))
        )
        Text(label, style = MaterialTheme.typography.labelSmall)
    }
}

@Composable
private fun DetalleDashboardDialog(
    detalle: DashboardDetalleInfo,
    onClose: () -> Unit,
    onShareText: () -> Unit,
    onShareImage: () -> Unit,
    onSharePdf: () -> Unit
) {
    androidx.compose.material3.AlertDialog(
        onDismissRequest = onClose,
        confirmButton = {
            androidx.compose.material3.TextButton(onClick = onShareText) {
                Text("Compartir")
            }
        },
        dismissButton = {
            Row(horizontalArrangement = Arrangement.spacedBy(8.dp)) {
                androidx.compose.material3.TextButton(onClick = onShareImage) {
                    Text("Imagen")
                }
                androidx.compose.material3.TextButton(onClick = onSharePdf) {
                    Text("PDF")
                }
                androidx.compose.material3.TextButton(onClick = onClose) {
                    Text("Cerrar")
                }
            }
        },
        title = { Text(detalle.title) },
        text = { Text(detalle.message) }
    )
}

private data class DashboardDetalleInfo(
    val title: String,
    val message: String
)

private fun DashboardDetalle.toDetalleInfo(state: com.prestamos.app.ui.model.DashboardResumen): DashboardDetalleInfo {
    return when (this) {
        DashboardDetalle.CAPITAL -> {
            val resumen = state.prestamosActivosDetalle
            val totalPrestamos = resumen.size
            val top = resumen.take(5).joinToString("\n") {
                "• ${it.cliente} | Prest. ${it.montoPrestado.toMoney(it.moneda)} | Saldo ${it.saldoPendiente.toMoney(it.moneda)} | Cuotas pend. ${it.cuotasPendientes}/${it.totalCuotas}"
            }.ifBlank { "No hay préstamos activos." }
            DashboardDetalleInfo(
                title = "Capital prestado",
                message = "Préstamos activos: $totalPrestamos\nTotal colocado: ${state.capitalPrestado.toMoney(state.monedaReferencial)}\n\n$top"
            )
        }

        DashboardDetalle.PENDIENTE -> {
            val resumen = state.cuotasPendientesDetalle
            val top = resumen.take(8).joinToString("\n") {
                "• ${it.cliente} | Préstamo #${it.idPrestamo} | Cuota ${it.numeroCuota} | Vence ${it.fechaVencimiento.toDateString()} | Saldo ${it.saldoPendiente.toMoney(it.moneda)}"
            }.ifBlank { "No hay cuotas pendientes." }
            DashboardDetalleInfo(
                title = "Saldo pendiente",
                message = "Cuotas pendientes: ${resumen.size}\nSaldo total pendiente: ${state.saldoPendiente.toMoney(state.monedaReferencial)}\n\n$top"
            )
        }

        DashboardDetalle.COBRADO_HOY -> {
            val resumen = state.pagosHoyDetalle
            val top = resumen.take(8).joinToString("\n") {
                "• ${it.cliente} | Préstamo #${it.idPrestamo} | Cuota ${it.numeroCuota} | ${it.fechaPago.toDateString()} | Abono ${it.montoAbono.toMoney(it.moneda)}"
            }.ifBlank { "No se registraron pagos hoy." }
            DashboardDetalleInfo(
                title = "Cobrado hoy",
                message = "Pagos de hoy: ${resumen.size}\nTotal cobrado hoy: ${state.cobradoHoy.toMoney(state.monedaReferencial)}\n\n$top"
            )
        }

        DashboardDetalle.VENCIDAS -> {
            val resumen = state.cuotasVencidasDetalle
            val top = resumen.take(8).joinToString("\n") {
                "• ${it.cliente} | Préstamo #${it.idPrestamo} | Cuota ${it.numeroCuota} | Vence ${it.fechaVencimiento.toDateString()} | Saldo ${it.saldoPendiente.toMoney(it.moneda)}"
            }.ifBlank { "No hay cuotas vencidas." }
            DashboardDetalleInfo(
                title = "Cuotas vencidas",
                message = "Cuotas vencidas: ${state.cuotasVencidas}\n\n$top"
            )
        }

        DashboardDetalle.CUOTAS_PAGADAS -> DashboardDetalleInfo(
            title = "Cuotas pagadas",
            message = "Cuotas pagadas: ${state.estadoCuotas["Pagadas"] ?: 0}\nSe consideran canceladas al 100%."
        )

        DashboardDetalle.CUOTAS_PENDIENTES -> {
            val count = state.estadoCuotas["Pendientes"] ?: 0
            val top = state.cuotasPendientesDetalle.take(6).joinToString("\n") {
                "• ${it.cliente} | Cuota ${it.numeroCuota} | Saldo ${it.saldoPendiente.toMoney(it.moneda)}"
            }.ifBlank { "No hay cuotas pendientes." }
            DashboardDetalleInfo(
                title = "Cuotas pendientes",
                message = "Cuotas en estado pendiente: $count\n\n$top"
            )
        }

        DashboardDetalle.CUOTAS_PARCIALES -> {
            val parciales = state.cuotasPendientesDetalle.filter { it.estado.name == "PARCIAL" }
            val top = parciales.take(6).joinToString("\n") {
                "• ${it.cliente} | Cuota ${it.numeroCuota} | Saldo ${it.saldoPendiente.toMoney(it.moneda)}"
            }.ifBlank { "No hay cuotas parciales." }
            DashboardDetalleInfo(
                title = "Cuotas parciales",
                message = "Cuotas parciales: ${state.estadoCuotas["Parciales"] ?: 0}\n\n$top"
            )
        }

        DashboardDetalle.CUOTAS_VENCIDAS -> {
            val top = state.cuotasVencidasDetalle.take(6).joinToString("\n") {
                "• ${it.cliente} | Cuota ${it.numeroCuota} | Vence ${it.fechaVencimiento.toDateString()} | Saldo ${it.saldoPendiente.toMoney(it.moneda)}"
            }.ifBlank { "No hay cuotas vencidas." }
            DashboardDetalleInfo(
                title = "Estado vencidas",
                message = "Cuotas en estado vencido: ${state.estadoCuotas["Vencidas"] ?: 0}\n\n$top"
            )
        }
    }
}

private fun compartirTextoDetalle(context: android.content.Context, detalle: DashboardDetalleInfo) {
    val sendIntent = Intent(Intent.ACTION_SEND).apply {
        type = "text/plain"
        putExtra(Intent.EXTRA_SUBJECT, "Detalle dashboard: ${detalle.title}")
        putExtra(Intent.EXTRA_TEXT, "${detalle.title}\n${detalle.message}")
    }
    context.startActivity(Intent.createChooser(sendIntent, "Compartir detalle"))
}

private fun compartirArchivoDetalle(context: android.content.Context, file: File, mimeType: String) {
    val uri = FileProvider.getUriForFile(
        context,
        "${context.packageName}.fileprovider",
        file
    )
    val sendIntent = Intent(Intent.ACTION_SEND).apply {
        type = mimeType
        putExtra(Intent.EXTRA_STREAM, uri)
        putExtra(Intent.EXTRA_SUBJECT, "Detalle dashboard")
        putExtra(Intent.EXTRA_TEXT, "Detalle exportado desde dashboard")
        addFlags(Intent.FLAG_GRANT_READ_URI_PERMISSION)
    }
    context.startActivity(Intent.createChooser(sendIntent, "Compartir detalle"))
}
