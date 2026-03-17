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
import androidx.compose.foundation.layout.width
import androidx.compose.foundation.lazy.LazyColumn
import androidx.compose.foundation.lazy.items
import androidx.compose.foundation.shape.RoundedCornerShape
import androidx.compose.material3.Card
import androidx.compose.material3.MaterialTheme
import androidx.compose.material3.OutlinedTextField
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
import com.prestamos.app.ui.theme.AccentGold
import com.prestamos.app.ui.theme.PrimaryGreen
import com.prestamos.app.ui.theme.SecondaryGreen
import androidx.compose.ui.unit.dp
import androidx.core.content.FileProvider
import androidx.lifecycle.compose.collectAsStateWithLifecycle
import com.prestamos.app.ui.screen.export.createDashboardDetallePdf
import com.prestamos.app.ui.viewmodel.ActivationUiState
import com.prestamos.app.ui.viewmodel.DashboardViewModel
import com.prestamos.app.util.toDateString
import com.prestamos.app.util.toMoney
import java.io.File

private enum class DashboardDetalle {
    CAPITAL,
    HISTORIAL,
    CAPITAL_ACTIVO2,
    PENDIENTE,
    COBRADO_HOY,
    VENCIDAS,
    CUOTAS_PAGADAS,
    CUOTAS_PENDIENTES,
    CUOTAS_PARCIALES,
    CUOTAS_VENCIDAS,
    GANANCIAS
}

@Composable
fun DashboardScreen(
    viewModel: DashboardViewModel,
    isDarkMode: Boolean,
    onToggleDarkMode: (Boolean) -> Unit,
    activationUiState: ActivationUiState,
    onActivationKeyChanged: (String) -> Unit,
    onActivateLicense: () -> Unit,
    onRefreshLicenseStatus: () -> Unit
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
                    Row(verticalAlignment = Alignment.CenterVertically, horizontalArrangement = Arrangement.spacedBy(6.dp)) {
                        val trialTexto = when {
                            activationUiState.status.licenseType.name == "TRIAL" && !activationUiState.status.trialExpired -> "Trial: ${activationUiState.status.trialDaysRemaining} día(s) restantes"
                            activationUiState.status.licenseType.name == "TRIAL" && activationUiState.status.trialExpired -> "Trial expirado"
                            activationUiState.status.licenseType.name == "ANUAL" -> "Licencia ANUAL activa"
                            else -> "Licencia FULL activa"
                        }
                        val trialColor = if (activationUiState.status.licenseType.name == "TRIAL") Color.Red else MaterialTheme.colorScheme.primary
                        Text(trialTexto, color = trialColor, style = MaterialTheme.typography.bodyMedium)
                    }
                }
            }
        }

        item {
            val capitalPrestadoActivo2 = state.prestamosActivosDetalle.sumOf { it.montoPrestado }
            Text(
                text = "Historial de prestamos",
                color = MaterialTheme.colorScheme.primary,
                modifier = Modifier.clickable { detalleSeleccionado = DashboardDetalle.HISTORIAL }
            )
            Spacer(modifier = Modifier.height(8.dp))
            DashboardCard("Capital prestado activo", capitalPrestadoActivo2.toMoney(state.monedaReferencial), Modifier.fillMaxWidth(), highlightValue = true) {
                detalleSeleccionado = DashboardDetalle.CAPITAL_ACTIVO2
            }
            Spacer(modifier = Modifier.height(8.dp))
            Row(horizontalArrangement = Arrangement.spacedBy(8.dp), modifier = Modifier.fillMaxWidth()) {
                DashboardCard("Saldo pendiente", state.saldoPendiente.toMoney(state.monedaReferencial), Modifier.weight(1f)) {
                    detalleSeleccionado = DashboardDetalle.PENDIENTE
                }
                DashboardCard("Cobrado hoy", state.cobradoHoy.toMoney(state.monedaReferencial), Modifier.weight(1f), highlightValue = true) {
                    detalleSeleccionado = DashboardDetalle.COBRADO_HOY
                }
            }
            Spacer(modifier = Modifier.height(8.dp))
            DashboardCard("Cuotas vencidas", state.cuotasVencidas.toString(), Modifier.fillMaxWidth()) {
                detalleSeleccionado = DashboardDetalle.VENCIDAS
            }
        }

        item {
            Text("Gráfico comparativo", style = MaterialTheme.typography.titleMedium)
            val capitalPrestadoActivo = state.prestamosActivosDetalle.sumOf { it.montoPrestado }
            val total = (capitalPrestadoActivo + state.cobradoHoy + state.saldoPendiente).coerceAtLeast(1.0)
            Row(modifier = Modifier.fillMaxWidth(), horizontalArrangement = Arrangement.spacedBy(8.dp)) {
                BarSegment("Prestado", (capitalPrestadoActivo / total).toFloat(), PrimaryGreen) { detalleSeleccionado = DashboardDetalle.CAPITAL }
                BarSegment("Cobrado", (state.cobradoHoy / total).toFloat(), AccentGold) { detalleSeleccionado = DashboardDetalle.COBRADO_HOY }
                BarSegment("Pendiente", (state.saldoPendiente / total).toFloat(), SecondaryGreen) { detalleSeleccionado = DashboardDetalle.PENDIENTE }
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
            Text("Ganancias de préstamos pagados", style = MaterialTheme.typography.titleMedium)
            DashboardCard(
                "Ganancia acumulada",
                state.gananciaAcumulada.toMoney(state.monedaReferencial),
                Modifier.fillMaxWidth()
            ) {
                detalleSeleccionado = DashboardDetalle.GANANCIAS
            }

            Spacer(Modifier.height(8.dp))
            if (state.gananciasPrestamosPagados.isEmpty()) {
                Text("No hay préstamos pagados aún")
            } else {
                val maxGanancia = state.gananciasPrestamosPagados.maxOf { it.ganancia }.coerceAtLeast(1.0)
                Column(verticalArrangement = Arrangement.spacedBy(8.dp)) {
                    state.gananciasPrestamosPagados.take(8).forEach { item ->
                        val ratio = (item.ganancia / maxGanancia).toFloat().coerceIn(0.1f, 1f)
                        Row(
                            modifier = Modifier.fillMaxWidth(),
                            verticalAlignment = Alignment.CenterVertically,
                            horizontalArrangement = Arrangement.spacedBy(8.dp)
                        ) {
                            Text("#${item.idPrestamo}", modifier = Modifier.width(44.dp), style = MaterialTheme.typography.labelMedium)
                            Spacer(
                                modifier = Modifier
                                    .height(20.dp)
                                    .fillMaxWidth(ratio)
                                    .background(PrimaryGreen, RoundedCornerShape(6.dp))
                            )
                        }
                        Text(
                            "${item.cliente}: ganancia ${item.ganancia.toMoney(item.moneda)}",
                            style = MaterialTheme.typography.bodySmall
                        )
                    }
                }
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
private fun DashboardCard(
    title: String,
    value: String,
    modifier: Modifier = Modifier,
    highlightValue: Boolean = false,
    onClick: () -> Unit
) {
    Card(
        modifier = modifier.clickable(onClick = onClick),
        colors = androidx.compose.material3.CardDefaults.cardColors(
            containerColor = MaterialTheme.colorScheme.secondaryContainer,
            contentColor = MaterialTheme.colorScheme.onSecondaryContainer
        )
    ) {
        Column(Modifier.padding(12.dp)) {
            Text(title, style = MaterialTheme.typography.labelLarge)
            Text(
                value,
                style = MaterialTheme.typography.titleLarge,
                color = if (highlightValue) MaterialTheme.colorScheme.tertiary else MaterialTheme.colorScheme.onSecondaryContainer
            )
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
            val totalActivo = resumen.sumOf { it.montoPrestado }
            val top = resumen.take(8).joinToString("\n") {
                "• ${it.cliente} | Préstamo #${it.idPrestamo} | Capital ${it.montoPrestado.toMoney(it.moneda)} | Estado: ACTIVO"
            }.ifBlank { "No hay préstamos activos." }
            DashboardDetalleInfo(
                title = "Capital prestado",
                message = "Préstamos activos: $totalPrestamos\nTotal activo: ${totalActivo.toMoney(state.monedaReferencial)}\n\n$top"
            )
        }


        DashboardDetalle.HISTORIAL -> {
            val historial = state.prestamosCapitalDetalle
                .filter { it.cuotasPendientes == 0 }
                .sortedByDescending { it.idPrestamo }
            val totalCobrado = historial.sumOf { it.montoCobrado }
            val top = historial.take(20).joinToString("\n") {
                "• ${it.cliente} | Préstamo #${it.idPrestamo} | Capital ${it.montoPrestado.toMoney(it.moneda)} | Cobrado ${it.montoCobrado.toMoney(it.moneda)} | Estado: PAGADO"
            }.ifBlank { "No hay préstamos cerrados todavía." }
            DashboardDetalleInfo(
                title = "Historial de préstamos",
                message = "Préstamos no activos o pagados: ${historial.size}\nTotal cobrado: ${totalCobrado.toMoney(state.monedaReferencial)}\n\n$top"
            )
        }

        DashboardDetalle.CAPITAL_ACTIVO2 -> {
            val resumen = state.prestamosActivosDetalle
            val total = resumen.sumOf { it.montoPrestado }
            val top = resumen.take(5).joinToString("\n") {
                "• ${it.cliente} | Préstamo #${it.idPrestamo} | Capital ${it.montoPrestado.toMoney(it.moneda)}"
            }.ifBlank { "No hay préstamos activos." }
            DashboardDetalleInfo(
                title = "Capital prestado activo",
                message = "Capital de préstamos activos: ${total.toMoney(state.monedaReferencial)}\n\n$top"
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

        DashboardDetalle.GANANCIAS -> {
            val top = state.gananciasPrestamosPagados.take(12).joinToString("\n") {
                "• ${it.cliente} | Préstamo #${it.idPrestamo} | Prestado ${it.montoPrestado.toMoney(it.moneda)} | Cobrado ${it.montoCobrado.toMoney(it.moneda)} | Ganancia ${it.ganancia.toMoney(it.moneda)}"
            }.ifBlank { "No hay préstamos pagados." }
            DashboardDetalleInfo(
                title = "Ganancias por préstamos pagados",
                message = "Préstamos pagados: ${state.gananciasPrestamosPagados.size}\nGanancia acumulada: ${state.gananciaAcumulada.toMoney(state.monedaReferencial)}\n\n$top"
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
