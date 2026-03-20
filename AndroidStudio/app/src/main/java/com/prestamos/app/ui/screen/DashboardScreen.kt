package com.prestamos.app.ui.screen

import android.content.Intent
import android.widget.Toast
import androidx.compose.foundation.background
import androidx.compose.foundation.clickable
import androidx.compose.foundation.layout.Box
import androidx.compose.foundation.layout.Arrangement
import androidx.compose.foundation.layout.Column
import androidx.compose.foundation.layout.Row
import androidx.compose.foundation.layout.Spacer
import androidx.compose.foundation.layout.fillMaxSize
import androidx.compose.foundation.layout.fillMaxWidth
import androidx.compose.foundation.layout.height
import androidx.compose.foundation.layout.heightIn
import androidx.compose.foundation.layout.padding
import androidx.compose.foundation.layout.width
import androidx.compose.foundation.rememberScrollState
import androidx.compose.foundation.lazy.LazyColumn
import androidx.compose.foundation.lazy.items
import androidx.compose.foundation.shape.RoundedCornerShape
import androidx.compose.foundation.verticalScroll
import androidx.compose.material3.Card
import androidx.compose.material3.CardDefaults
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
import androidx.compose.ui.text.style.TextOverflow
import com.prestamos.app.ui.theme.AccentGold
import com.prestamos.app.ui.theme.PrimaryGreen
import com.prestamos.app.ui.theme.SecondaryGreen
import androidx.compose.ui.unit.dp
import androidx.core.content.FileProvider
import androidx.lifecycle.compose.collectAsStateWithLifecycle
import com.prestamos.app.data.config.InitialSetupPreferences
import com.prestamos.app.data.local.entity.Moneda
import com.prestamos.app.ui.screen.export.createDashboardDetallePdf
import com.prestamos.app.ui.viewmodel.ActivationUiState
import com.prestamos.app.ui.viewmodel.DashboardViewModel
import com.prestamos.app.util.toDateString
import com.prestamos.app.util.toMoney
import java.io.File
import java.time.Instant
import java.time.LocalDate
import java.time.ZoneId
import java.time.temporal.ChronoUnit
import java.util.Locale

private enum class DashboardDetalle {
    CAPITAL,
    HISTORIAL,
    CAPITAL_ACTIVO2,
    PENDIENTE,
    COBRADO_HOY,
    COBRADO_ACTIVO,
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
    val setupPrefs = remember { InitialSetupPreferences(context) }
    val businessName = remember { setupPrefs.getBusinessName().trim() }
    val visibleCurrencies = remember {
        resolveVisibleCurrencies(
            setupPrefs.getMainCurrencyCode(),
            setupPrefs.getSecondaryCurrencyCode()
        )
    }
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
                    if (businessName.isNotBlank()) {
                        Text(
                            businessName,
                            style = MaterialTheme.typography.titleMedium,
                            color = MaterialTheme.colorScheme.primary
                        )
                    }
                    Text("Fecha: ${System.currentTimeMillis().toDateString()}")
                    Row(verticalAlignment = Alignment.CenterVertically, horizontalArrangement = Arrangement.spacedBy(6.dp)) {
                        val trialTexto = when {
                            activationUiState.status.licenseType.name == "TRIAL" && !activationUiState.status.trialExpired -> "Trial: ${activationUiState.status.trialDaysRemaining} dia(s) restantes"
                            activationUiState.status.licenseType.name == "TRIAL" && activationUiState.status.trialExpired -> "Trial expirado"
                            activationUiState.status.licenseType.name == "MENSUAL" -> "Licencia MENSUAL activa"
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
            val capitalPrestadoActivo2 = state.prestamosActivosDetalle.sumByCurrency { it.montoPrestado }
            val prestadoActivoConInteres = state.prestamosActivosDetalle.sumByCurrency { it.montoTotalConInteres }
            Text(
                text = "Historial de prestamos",
                color = MaterialTheme.colorScheme.primary,
                modifier = Modifier.clickable { detalleSeleccionado = DashboardDetalle.HISTORIAL }
            )
            Spacer(modifier = Modifier.height(8.dp))
            DashboardMoneyCard(
                title = "Capital prestado activo",
                totals = capitalPrestadoActivo2,
                visibleCurrencies = visibleCurrencies,
                modifier = Modifier.fillMaxWidth(),
                highlightValue = true,
                valueColor = MaterialTheme.colorScheme.onSecondaryContainer
            ) {
                detalleSeleccionado = DashboardDetalle.CAPITAL_ACTIVO2
            }
            Spacer(modifier = Modifier.height(8.dp))
            DashboardMoneyCard(
                title = "Prestado activo + intereses",
                totals = prestadoActivoConInteres,
                visibleCurrencies = visibleCurrencies,
                modifier = Modifier.fillMaxWidth(),
                highlightValue = true,
                valueColor = MaterialTheme.colorScheme.onSecondaryContainer
            ) {
                detalleSeleccionado = DashboardDetalle.CAPITAL
            }
            Spacer(modifier = Modifier.height(8.dp))
            Row(horizontalArrangement = Arrangement.spacedBy(8.dp), modifier = Modifier.fillMaxWidth()) {
                DashboardMoneyCard(
                    title = "Saldo pendiente",
                    totals = state.cuotasPendientesDetalle.sumByCurrency { it.saldoPendiente },
                    visibleCurrencies = visibleCurrencies,
                    modifier = Modifier.weight(1f)
                ) {
                    detalleSeleccionado = DashboardDetalle.PENDIENTE
                }
                DashboardMoneyCard(
                    title = "Cobrado hoy",
                    totals = state.pagosHoyDetalle.sumByCurrency { it.montoAbono },
                    visibleCurrencies = visibleCurrencies,
                    modifier = Modifier.weight(1f),
                    highlightValue = true,
                    valueColor = MaterialTheme.colorScheme.onSecondaryContainer
                ) {
                    detalleSeleccionado = DashboardDetalle.COBRADO_HOY
                }
            }
            Spacer(modifier = Modifier.height(8.dp))
            DashboardCard("Cuotas vencidas", state.cuotasVencidas.toString(), Modifier.fillMaxWidth()) {
                detalleSeleccionado = DashboardDetalle.VENCIDAS
            }
        }

        item {
            Text("Grafico comparativo", style = MaterialTheme.typography.titleMedium)
            Column(verticalArrangement = Arrangement.spacedBy(10.dp)) {
                visibleCurrencies.forEach { moneda ->
                    val activosMoneda = state.prestamosActivosDetalle.filter { it.moneda == moneda }
                    val capitalPrestadoActivo = activosMoneda.sumOf { it.montoTotalConInteres }.coerceAtLeast(0.0)
                    val pendienteActivo = activosMoneda
                        .sumOf { it.saldoPendiente }
                        .coerceIn(0.0, capitalPrestadoActivo)
                    val cobradoActivo = (capitalPrestadoActivo - pendienteActivo).coerceAtLeast(0.0)
                    Card(
                        modifier = Modifier.fillMaxWidth(),
                        colors = CardDefaults.cardColors(
                            containerColor = MaterialTheme.colorScheme.secondaryContainer.copy(alpha = 0.55f),
                            contentColor = MaterialTheme.colorScheme.onSecondaryContainer
                        ),
                        shape = RoundedCornerShape(12.dp)
                    ) {
                        Column(
                            modifier = Modifier
                                .fillMaxWidth()
                                .padding(10.dp),
                            verticalArrangement = Arrangement.spacedBy(6.dp)
                        ) {
                            Text(
                                text = "Totales en ${moneda.displayName}",
                                style = MaterialTheme.typography.labelLarge
                            )
                            HorizontalMetricBar(
                                label = "Prestado activo + interes",
                                amount = capitalPrestadoActivo,
                                percent = 1f,
                                color = PrimaryGreen,
                                moneda = moneda,
                                onClick = { detalleSeleccionado = DashboardDetalle.CAPITAL }
                            )
                            HorizontalMetricBar(
                                label = "Cobrado activo + interes",
                                amount = cobradoActivo,
                                percent = if (capitalPrestadoActivo > 0.0) (cobradoActivo / capitalPrestadoActivo).toFloat() else 0f,
                                color = AccentGold,
                                moneda = moneda,
                                onClick = { detalleSeleccionado = DashboardDetalle.COBRADO_ACTIVO }
                            )
                            HorizontalMetricBar(
                                label = "Pendiente activo + interes",
                                amount = pendienteActivo,
                                percent = if (capitalPrestadoActivo > 0.0) (pendienteActivo / capitalPrestadoActivo).toFloat() else 0f,
                                color = SecondaryGreen,
                                moneda = moneda,
                                onClick = { detalleSeleccionado = DashboardDetalle.PENDIENTE }
                            )
                        }
                    }
                }
            }
        }

        item {
            Text("Ganancias de prestamos pagados", style = MaterialTheme.typography.titleMedium)
            DashboardMoneyCard(
                title = "Ganancia acumulada",
                totals = state.gananciasPrestamosPagados.sumByCurrency { it.ganancia },
                visibleCurrencies = visibleCurrencies,
                modifier = Modifier.fillMaxWidth()
            ) {
                detalleSeleccionado = DashboardDetalle.GANANCIAS
            }

            Spacer(Modifier.height(8.dp))
            if (state.gananciasPrestamosPagados.isEmpty()) {
                Text("No hay prestamos pagados aun")
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
            Text("Proximos vencimientos", style = MaterialTheme.typography.titleMedium)
            if (state.proximosVencimientos.isEmpty()) {
                Text("No hay Proximos vencimientos")
            }
        }
        items(state.proximosVencimientos) { cuota ->
            VencimientoCard(cuota)
        }

        item {
            Text("Ultimos pagos", style = MaterialTheme.typography.titleMedium)
            if (state.ultimosPagos.isEmpty()) {
                Text("No hay pagos registrados")
            }
        }
        items(state.ultimosPagos) { pago ->
            PagoRecienteCard(pago)
        }
    }

    detalleSeleccionado?.let {
        if (it == DashboardDetalle.CAPITAL_ACTIVO2 || it == DashboardDetalle.CAPITAL) {
            val capitalDetalle = it.toCapitalDetalleUi(state, visibleCurrencies)
            CapitalDetalleDialog(
                detalle = capitalDetalle,
                onClose = { detalleSeleccionado = null },
                onShareText = {
                    compartirTextoPlano(
                        context = context,
                        titulo = capitalDetalle.title,
                        detalle = capitalDetalle.toShareText()
                    )
                },
                onSharePdf = {
                    runCatching {
                        createDashboardDetallePdf(context, capitalDetalle.title, capitalDetalle.toShareText())
                    }.onSuccess { file ->
                        compartirArchivoDetalle(context, file, "application/pdf")
                    }.onFailure {
                        Toast.makeText(context, "No se pudo exportar PDF", Toast.LENGTH_SHORT).show()
                    }
                }
            )
        } else if (it == DashboardDetalle.PENDIENTE) {
            val pendienteDetalle = it.toPendienteDetalleUi(state, visibleCurrencies)
            PendienteDetalleDialog(
                detalle = pendienteDetalle,
                onClose = { detalleSeleccionado = null },
                onShareText = {
                    compartirTextoPlano(
                        context = context,
                        titulo = pendienteDetalle.title,
                        detalle = pendienteDetalle.toShareText()
                    )
                },
                onSharePdf = {
                    runCatching {
                        createDashboardDetallePdf(context, pendienteDetalle.title, pendienteDetalle.toShareText())
                    }.onSuccess { file ->
                        compartirArchivoDetalle(context, file, "application/pdf")
                    }.onFailure {
                        Toast.makeText(context, "No se pudo exportar PDF", Toast.LENGTH_SHORT).show()
                    }
                }
            )
        } else {
            val detalle = it.toDetalleInfo(state, visibleCurrencies)
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
}

@Composable
private fun DashboardCard(
    title: String,
    value: String,
    modifier: Modifier = Modifier,
    highlightValue: Boolean = false,
    valueColor: Color? = null,
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
                color = valueColor ?: if (highlightValue) {
                    MaterialTheme.colorScheme.tertiary
                } else {
                    MaterialTheme.colorScheme.onSecondaryContainer
                }
            )
        }
    }
}

@Composable
private fun HorizontalMetricBar(
    label: String,
    amount: Double,
    percent: Float,
    color: Color,
    moneda: com.prestamos.app.data.local.entity.Moneda,
    onClick: () -> Unit
) {
    val progress = percent.coerceIn(0f, 1f)
    val percentText = String.format(Locale.US, "%.1f%%", progress * 100f)

    Column(
        modifier = Modifier
            .fillMaxWidth()
            .clickable(onClick = onClick),
        verticalArrangement = Arrangement.spacedBy(4.dp)
    ) {
        Row(
            modifier = Modifier.fillMaxWidth(),
            horizontalArrangement = Arrangement.SpaceBetween
        ) {
            Text(label, style = MaterialTheme.typography.labelMedium)
            Text("$percentText  ${amount.toMoney(moneda)}", style = MaterialTheme.typography.labelMedium)
        }
        Box(
            modifier = Modifier
                .fillMaxWidth()
                .height(12.dp)
                .background(MaterialTheme.colorScheme.surfaceVariant, RoundedCornerShape(8.dp))
        ) {
            Box(
                modifier = Modifier
                    .fillMaxWidth(progress)
                    .height(12.dp)
                    .background(color, RoundedCornerShape(8.dp))
            )
        }
    }
}

private enum class UrgenciaVencimiento {
    VENCIDO,
    PRONTO,
    NORMAL
}

@Composable
private fun VencimientoCard(cuota: com.prestamos.app.ui.model.DashboardCuotaItem) {
    val iconAlert = "\u26A0\uFE0F"
    val iconSoon = "\u23F3"
    val iconPin = "\uD83D\uDCCC"
    val iconCal = "\uD83D\uDCC5"
    val iconMoney = "\uD83D\uDCB0"

    val today = LocalDate.now()
    val fechaVenc = Instant.ofEpochMilli(cuota.fechaVencimiento).atZone(ZoneId.systemDefault()).toLocalDate()
    val dias = ChronoUnit.DAYS.between(today, fechaVenc).toInt()
    val urgencia = when {
        dias < 0 -> UrgenciaVencimiento.VENCIDO
        dias in 0..3 -> UrgenciaVencimiento.PRONTO
        else -> UrgenciaVencimiento.NORMAL
    }
    val estadoTexto = when {
        dias < 0 -> "$iconAlert Vencido hace ${kotlin.math.abs(dias)} ${if (kotlin.math.abs(dias) == 1) "dia" else "dias"}"
        dias == 0 -> "$iconSoon Vence hoy"
        else -> "$iconSoon Vence en $dias ${if (dias == 1) "dia" else "dias"}"
    }
    val fondo = when (urgencia) {
        UrgenciaVencimiento.VENCIDO -> Color(0xFFFFE6E6)
        UrgenciaVencimiento.PRONTO -> Color(0xFFFFF0DB)
        UrgenciaVencimiento.NORMAL -> Color(0xFFEAF6EA)
    }

    Card(
        modifier = Modifier.fillMaxWidth(),
        colors = CardDefaults.cardColors(
            containerColor = fondo,
            contentColor = Color(0xFF1F1F1F)
        )
    ) {
        Column(
            modifier = Modifier.padding(horizontal = 12.dp, vertical = 10.dp),
            verticalArrangement = Arrangement.spacedBy(4.dp)
        ) {
            Text("${cuota.cliente} - $estadoTexto", style = MaterialTheme.typography.titleSmall, color = Color(0xFF1F1F1F))
            Text("$iconPin Cuota ${cuota.numeroCuota} \u2022 Prestamo #${cuota.idPrestamo}", style = MaterialTheme.typography.bodyMedium, color = Color(0xFF1F1F1F))
            Text("$iconCal ${cuota.fechaVencimiento.toDateString()}", style = MaterialTheme.typography.bodyMedium, color = Color(0xFF1F1F1F))
            Text(
                "$iconMoney ${cuota.saldoPendiente.toMoney(cuota.moneda)} ${cuota.moneda.displayName}",
                style = MaterialTheme.typography.bodyMedium,
                color = Color(0xFF1F1F1F)
            )
        }
    }
}

@Composable
private fun PagoRecienteCard(pago: com.prestamos.app.ui.model.DashboardPagoItem) {
    val iconPin = "\uD83D\uDCCC"
    val iconCal = "\uD83D\uDCC5"
    val iconMoney = "\uD83D\uDCB0"

    Card(
        modifier = Modifier.fillMaxWidth(),
        colors = CardDefaults.cardColors(
            containerColor = MaterialTheme.colorScheme.secondaryContainer.copy(alpha = 0.55f),
            contentColor = MaterialTheme.colorScheme.onSecondaryContainer
        )
    ) {
        Column(
            modifier = Modifier.padding(horizontal = 12.dp, vertical = 10.dp),
            verticalArrangement = Arrangement.spacedBy(4.dp)
        ) {
            Text(pago.cliente, style = MaterialTheme.typography.titleSmall)
            Text("$iconPin Cuota ${pago.numeroCuota} \u2022 Prestamo #${pago.idPrestamo}", style = MaterialTheme.typography.bodyMedium)
            Text("$iconCal ${pago.fechaPago.toDateString()}", style = MaterialTheme.typography.bodyMedium)
            Text("$iconMoney ${pago.montoAbono.toMoney(pago.moneda)} ${pago.moneda.displayName}", style = MaterialTheme.typography.bodyMedium)
        }
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

private data class CapitalDetalleItem(
    val cliente: String,
    val idPrestamo: Long,
    val moneda: Moneda,
    val monto: Double
)

private data class CapitalDetalleUi(
    val title: String,
    val totalsByCurrency: Map<Moneda, Double>,
    val visibleCurrencies: List<Moneda>,
    val items: List<CapitalDetalleItem>
)

private data class PendienteDetalleItem(
    val cliente: String,
    val idPrestamo: Long,
    val numeroCuota: Int,
    val fechaVencimiento: Long,
    val saldoPendiente: Double,
    val moneda: Moneda
)

private data class PendienteDetalleUi(
    val title: String,
    val totalCuotas: Int,
    val totalsByCurrency: Map<Moneda, Double>,
    val visibleCurrencies: List<Moneda>,
    val items: List<PendienteDetalleItem>
)

private fun CapitalDetalleUi.toShareText(): String = buildString {
    appendLine(title)
    visibleCurrencies.forEach { moneda ->
        appendLine("Total ${moneda.displayName}: ${(totalsByCurrency[moneda] ?: 0.0).toMoney(moneda)}")
    }
    appendLine()
    appendLine("Detalle de prestamos")
    if (items.isEmpty()) {
        appendLine("No hay prestamos activos.")
    } else {
        items.forEach { item ->
            appendLine("${item.cliente} | Prestamo #${item.idPrestamo} | ${item.monto.toMoney(item.moneda)} ${item.moneda.displayName}")
        }
    }
}

private fun PendienteDetalleUi.toShareText(): String = buildString {
    appendLine(title)
    appendLine("Cuotas pendientes: $totalCuotas")
    visibleCurrencies.forEach { moneda ->
        appendLine("Total pendiente ${moneda.displayName}: ${(totalsByCurrency[moneda] ?: 0.0).toMoney(moneda)}")
    }
    appendLine()
    appendLine("Detalle")
    if (items.isEmpty()) {
        appendLine("No hay cuotas pendientes.")
    } else {
        items.forEach { item ->
            appendLine("${item.cliente} | Prestamo #${item.idPrestamo} | Cuota ${item.numeroCuota} | Vence ${item.fechaVencimiento.toDateString()} | Saldo ${item.saldoPendiente.toMoney(item.moneda)} ${item.moneda.displayName}")
        }
    }
}

@Composable
private fun CapitalDetalleDialog(
    detalle: CapitalDetalleUi,
    onClose: () -> Unit,
    onShareText: () -> Unit,
    onSharePdf: () -> Unit
) {
    androidx.compose.material3.AlertDialog(
        onDismissRequest = onClose,
        confirmButton = {
            Row(horizontalArrangement = Arrangement.spacedBy(8.dp)) {
                androidx.compose.material3.TextButton(onClick = onSharePdf) {
                    Text("PDF")
                }
                androidx.compose.material3.TextButton(onClick = onClose) {
                    Text("Cerrar")
                }
                androidx.compose.material3.TextButton(onClick = onShareText) {
                    Text("Compartir")
                }
            }
        },
        title = { Text(detalle.title) },
        text = {
            Column(
                verticalArrangement = Arrangement.spacedBy(8.dp),
                modifier = Modifier
                    .heightIn(max = 420.dp)
                    .verticalScroll(rememberScrollState())
            ) {
                detalle.visibleCurrencies.forEach { moneda ->
                    Card(
                        modifier = Modifier.fillMaxWidth(),
                        colors = CardDefaults.cardColors(
                            containerColor = MaterialTheme.colorScheme.surfaceVariant.copy(alpha = 0.55f)
                        )
                    ) {
                        Row(
                            modifier = Modifier
                                .fillMaxWidth()
                                .padding(horizontal = 8.dp, vertical = 6.dp),
                            horizontalArrangement = Arrangement.SpaceBetween
                        ) {
                            Text("💰 Total ${moneda.displayName}", style = MaterialTheme.typography.bodyMedium)
                            Text((detalle.totalsByCurrency[moneda] ?: 0.0).toMoney(moneda), style = MaterialTheme.typography.bodyMedium)
                        }
                    }
                }
                Text("📄 Detalle de prestamos", style = MaterialTheme.typography.labelLarge)
                if (detalle.items.isEmpty()) {
                    Text("No hay prestamos activos.")
                } else {
                    Column(verticalArrangement = Arrangement.spacedBy(6.dp)) {
                        detalle.items.forEach { item ->
                            Card(
                                modifier = Modifier.fillMaxWidth(),
                                colors = CardDefaults.cardColors(
                                    containerColor = MaterialTheme.colorScheme.secondaryContainer.copy(alpha = 0.6f)
                                )
                            ) {
                                Column(
                                    modifier = Modifier.padding(8.dp),
                                    verticalArrangement = Arrangement.spacedBy(2.dp)
                                ) {
                                    Text("👤 ${item.cliente}", style = MaterialTheme.typography.bodyMedium)
                                    Text("💰 ${item.monto.toMoney(item.moneda)} ${item.moneda.displayName}", style = MaterialTheme.typography.bodySmall)
                                    Text("📄 Prestamo #${item.idPrestamo}", style = MaterialTheme.typography.bodySmall)
                                }
                            }
                        }
                    }
                }
            }
        }
    )
}

@Composable
private fun PendienteDetalleDialog(
    detalle: PendienteDetalleUi,
    onClose: () -> Unit,
    onShareText: () -> Unit,
    onSharePdf: () -> Unit
) {
    androidx.compose.material3.AlertDialog(
        onDismissRequest = onClose,
        confirmButton = {
            Row(horizontalArrangement = Arrangement.spacedBy(8.dp)) {
                androidx.compose.material3.TextButton(onClick = onSharePdf) {
                    Text("PDF")
                }
                androidx.compose.material3.TextButton(onClick = onClose) {
                    Text("Cerrar")
                }
                androidx.compose.material3.TextButton(onClick = onShareText) {
                    Text("Compartir")
                }
            }
        },
        title = { Text(detalle.title) },
        text = {
            Column(
                verticalArrangement = Arrangement.spacedBy(8.dp),
                modifier = Modifier
                    .heightIn(max = 420.dp)
                    .verticalScroll(rememberScrollState())
            ) {
                Text("\uD83D\uDCCC Resumen", style = MaterialTheme.typography.labelLarge)
                Card(
                    modifier = Modifier.fillMaxWidth(),
                    colors = CardDefaults.cardColors(
                        containerColor = MaterialTheme.colorScheme.surfaceVariant.copy(alpha = 0.55f)
                    )
                ) {
                    Text(
                        text = "\uD83D\uDD22 Cuotas pendientes: ${detalle.totalCuotas}",
                        modifier = Modifier.padding(horizontal = 8.dp, vertical = 6.dp),
                        style = MaterialTheme.typography.bodyMedium
                    )
                }
                detalle.visibleCurrencies.forEach { moneda ->
                    Card(
                        modifier = Modifier.fillMaxWidth(),
                        colors = CardDefaults.cardColors(
                            containerColor = MaterialTheme.colorScheme.surfaceVariant.copy(alpha = 0.55f)
                        )
                    ) {
                        Text(
                            text = "\uD83D\uDCB0 Total pendiente ${moneda.displayName}: ${(detalle.totalsByCurrency[moneda] ?: 0.0).toMoney(moneda)}",
                            modifier = Modifier.padding(horizontal = 8.dp, vertical = 6.dp),
                            style = MaterialTheme.typography.bodyMedium
                        )
                    }
                }
                Text("\uD83D\uDCCB Detalle", style = MaterialTheme.typography.labelLarge)
                if (detalle.items.isEmpty()) {
                    Text("No hay cuotas pendientes.")
                } else {
                    Column(verticalArrangement = Arrangement.spacedBy(6.dp)) {
                        detalle.items.forEach { item ->
                            Card(
                                modifier = Modifier.fillMaxWidth(),
                                colors = CardDefaults.cardColors(
                                    containerColor = MaterialTheme.colorScheme.secondaryContainer.copy(alpha = 0.6f)
                                )
                            ) {
                                Column(
                                    modifier = Modifier.padding(8.dp),
                                    verticalArrangement = Arrangement.spacedBy(2.dp)
                                ) {
                                    Text("\uD83D\uDC64 ${item.cliente}", style = MaterialTheme.typography.bodyMedium)
                                    Text("\uD83D\uDCC4 Prestamo #${item.idPrestamo} \u00B7 Cuota ${item.numeroCuota}", style = MaterialTheme.typography.bodySmall)
                                    Text("\uD83D\uDCC5 Vence ${item.fechaVencimiento.toDateString()}", style = MaterialTheme.typography.bodySmall)
                                    Text("\uD83D\uDCB0 Saldo ${item.saldoPendiente.toMoney(item.moneda)} ${item.moneda.displayName}", style = MaterialTheme.typography.bodySmall)
                                }
                            }
                        }
                    }
                }
            }
        }
    )
}

private fun DashboardDetalle.toDetalleInfo(
    state: com.prestamos.app.ui.model.DashboardResumen,
    visibleCurrencies: List<Moneda>
): DashboardDetalleInfo {
    return when (this) {
        DashboardDetalle.CAPITAL -> {
            val resumen = state.prestamosActivosDetalle
            val totalPrestamos = resumen.size
            val totalActivo = resumen.sumByCurrency { it.montoTotalConInteres }
            val top = resumen.take(8).joinToString("\n") {
                "- ${it.cliente} | Prestamo #${it.idPrestamo} | Prestado c/interes ${it.montoTotalConInteres.toMoney(it.moneda)} | Estado: ACTIVO"
            }.ifBlank { "No hay prestamos activos." }
            DashboardDetalleInfo(
                title = "Prestado activo + interes",
                message = "Prestamos activos: $totalPrestamos\n${totalActivo.toTotalsText(visibleCurrencies)}\n\n$top"
            )
        }

        DashboardDetalle.HISTORIAL -> {
            val historial = state.prestamosCapitalDetalle
                .filter { it.cuotasPendientes == 0 }
                .sortedByDescending { it.idPrestamo }
            val totalCobrado = historial.sumByCurrency { it.montoCobrado }
            val totalGanado = historial.sumByCurrency { it.montoCobrado - it.montoPrestado }
            val top = historial.take(20).joinToString("\n") {
                "- ${it.cliente} | Prestamo #${it.idPrestamo} | Capital ${it.montoPrestado.toMoney(it.moneda)} | Cobrado ${it.montoCobrado.toMoney(it.moneda)} | Estado: PAGADO"
            }.ifBlank { "No hay prestamos cerrados todavia." }
            DashboardDetalleInfo(
                title = "Historial de prestamos",
                message = "Prestamos no activos o pagados: ${historial.size}\nTotal cobrado:\n${totalCobrado.toTotalsText(visibleCurrencies)}\nTotal ganado:\n${totalGanado.toTotalsText(visibleCurrencies)}\n\n$top"
            )
        }

        DashboardDetalle.CAPITAL_ACTIVO2 -> {
            val resumen = state.prestamosActivosDetalle
            val total = resumen.sumByCurrency { it.montoPrestado }
            val top = resumen.take(5).joinToString("\n") {
                "- ${it.cliente} | Prestamo #${it.idPrestamo} | Capital ${it.montoPrestado.toMoney(it.moneda)}"
            }.ifBlank { "No hay prestamos activos." }
            DashboardDetalleInfo(
                title = "Capital prestado activo",
                message = "Capital de prestamos activos:\n${total.toTotalsText(visibleCurrencies)}\n\n$top"
            )
        }

        DashboardDetalle.PENDIENTE -> {
            val resumen = state.cuotasPendientesDetalle
            val top = resumen.take(8).joinToString("\n") {
                "- ${it.cliente} | Prestamo #${it.idPrestamo} | Cuota ${it.numeroCuota} | Vence ${it.fechaVencimiento.toDateString()} | Saldo ${it.saldoPendiente.toMoney(it.moneda)}"
            }.ifBlank { "No hay cuotas pendientes." }
            DashboardDetalleInfo(
                title = "Saldo pendiente",
                message = "Cuotas pendientes: ${resumen.size}\nSaldo total pendiente:\n${resumen.sumByCurrency { it.saldoPendiente }.toTotalsText(visibleCurrencies)}\n\n$top"
            )
        }

        DashboardDetalle.COBRADO_HOY -> {
            val resumen = state.pagosHoyDetalle
            val top = resumen.take(8).joinToString("\n") {
                "- ${it.cliente} | Prestamo #${it.idPrestamo} | Cuota ${it.numeroCuota} | ${it.fechaPago.toDateString()} | Abono ${it.montoAbono.toMoney(it.moneda)}"
            }.ifBlank { "No se registraron pagos hoy." }
            DashboardDetalleInfo(
                title = "Cobrado hoy",
                message = "Pagos de hoy: ${resumen.size}\nTotal cobrado hoy:\n${resumen.sumByCurrency { it.montoAbono }.toTotalsText(visibleCurrencies)}\n\n$top"
            )
        }

        DashboardDetalle.COBRADO_ACTIVO -> {
            val resumen = state.prestamosActivosDetalle
            val totalCobradoActivo = resumen.sumByCurrency { (it.montoTotalConInteres - it.saldoPendiente).coerceAtLeast(0.0) }
            val top = resumen.take(12).joinToString("\n") {
                val cobradoPrestamo = (it.montoTotalConInteres - it.saldoPendiente).coerceAtLeast(0.0)
                "- ${it.cliente} | Prestamo #${it.idPrestamo} | Cobrado c/interes ${cobradoPrestamo.toMoney(it.moneda)} | Pendiente c/interes ${it.saldoPendiente.toMoney(it.moneda)}"
            }.ifBlank { "No hay prestamos activos." }
            DashboardDetalleInfo(
                title = "Cobrado activo",
                message = "Prestamos activos: ${resumen.size}\nTotal cobrado activo c/interes:\n${totalCobradoActivo.toTotalsText(visibleCurrencies)}\n\n$top"
            )
        }

        DashboardDetalle.VENCIDAS -> {
            val resumen = state.cuotasVencidasDetalle
            val top = resumen.take(8).joinToString("\n") {
                "- ${it.cliente} | Prestamo #${it.idPrestamo} | Cuota ${it.numeroCuota} | Vence ${it.fechaVencimiento.toDateString()} | Saldo ${it.saldoPendiente.toMoney(it.moneda)}"
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
                "- ${it.cliente} | Cuota ${it.numeroCuota} | Saldo ${it.saldoPendiente.toMoney(it.moneda)}"
            }.ifBlank { "No hay cuotas pendientes." }
            DashboardDetalleInfo(
                title = "Cuotas pendientes",
                message = "Cuotas en estado pendiente: $count\n\n$top"
            )
        }

        DashboardDetalle.CUOTAS_PARCIALES -> {
            val parciales = state.cuotasPendientesDetalle.filter { it.estado.name == "PARCIAL" }
            val top = parciales.take(6).joinToString("\n") {
                "- ${it.cliente} | Cuota ${it.numeroCuota} | Saldo ${it.saldoPendiente.toMoney(it.moneda)}"
            }.ifBlank { "No hay cuotas parciales." }
            DashboardDetalleInfo(
                title = "Cuotas parciales",
                message = "Cuotas parciales: ${state.estadoCuotas["Parciales"] ?: 0}\n\n$top"
            )
        }

        DashboardDetalle.CUOTAS_VENCIDAS -> {
            val top = state.cuotasVencidasDetalle.take(6).joinToString("\n") {
                "- ${it.cliente} | Cuota ${it.numeroCuota} | Vence ${it.fechaVencimiento.toDateString()} | Saldo ${it.saldoPendiente.toMoney(it.moneda)}"
            }.ifBlank { "No hay cuotas vencidas." }
            DashboardDetalleInfo(
                title = "Estado vencidas",
                message = "Cuotas en estado vencido: ${state.estadoCuotas["Vencidas"] ?: 0}\n\n$top"
            )
        }

        DashboardDetalle.GANANCIAS -> {
            val gananciaPorMoneda = state.gananciasPrestamosPagados.sumByCurrency { it.ganancia }
            val top = state.gananciasPrestamosPagados.take(12).joinToString("\n") {
                "- ${it.cliente} | Prestamo #${it.idPrestamo} | Prestado ${it.montoPrestado.toMoney(it.moneda)} | Cobrado ${it.montoCobrado.toMoney(it.moneda)} | Ganancia ${it.ganancia.toMoney(it.moneda)}"
            }.ifBlank { "No hay prestamos pagados." }
            DashboardDetalleInfo(
                title = "Ganancias por prestamos pagados",
                message = "Prestamos pagados: ${state.gananciasPrestamosPagados.size}\nGanancia acumulada:\n${gananciaPorMoneda.toTotalsText(visibleCurrencies)}\n\n$top"
            )
        }
    }
}

private fun <T> List<T>.sumByCurrency(
    currencySelector: (T) -> Moneda = { item ->
        when (item) {
            is com.prestamos.app.ui.model.DashboardPrestamoDetalleItem -> item.moneda
            is com.prestamos.app.ui.model.DashboardCuotaDetalleItem -> item.moneda
            is com.prestamos.app.ui.model.DashboardPagoItem -> item.moneda
            is com.prestamos.app.ui.model.DashboardGananciaPrestamoItem -> item.moneda
            else -> Moneda.SOLES
        }
    },
    amountSelector: (T) -> Double
): Map<Moneda, Double> {
    if (isEmpty()) return mapOf(Moneda.SOLES to 0.0)
    return groupBy(currencySelector)
        .mapValues { (_, values) -> values.sumOf(amountSelector) }
}

private fun Map<Moneda, Double>.toTotalsText(visibleCurrencies: List<Moneda>): String {
    val ordered = visibleCurrencies.ifEmpty { listOf(Moneda.SOLES) }
    return ordered.joinToString("\n") { moneda ->
        val total = this[moneda] ?: 0.0
        val label = "Totales en ${moneda.displayName}"
        "$label: ${total.toMoney(moneda)}"
    }
}

private fun DashboardDetalle.toCapitalDetalleUi(
    state: com.prestamos.app.ui.model.DashboardResumen,
    visibleCurrencies: List<Moneda>
): CapitalDetalleUi {
    val items = state.prestamosActivosDetalle.map { detalle ->
        CapitalDetalleItem(
            cliente = detalle.cliente,
            idPrestamo = detalle.idPrestamo,
            moneda = detalle.moneda,
            monto = if (this == DashboardDetalle.CAPITAL) detalle.montoTotalConInteres else detalle.montoPrestado
        )
    }.sortedByDescending { it.idPrestamo }

    val totals = items
        .groupBy { it.moneda }
        .mapValues { (_, values) -> values.sumOf { it.monto } }

    return CapitalDetalleUi(
        title = if (this == DashboardDetalle.CAPITAL) "Capital prestado activo + intereses" else "Capital prestado activo",
        totalsByCurrency = totals,
        visibleCurrencies = visibleCurrencies,
        items = items
    )
}

private fun DashboardDetalle.toPendienteDetalleUi(
    state: com.prestamos.app.ui.model.DashboardResumen,
    visibleCurrencies: List<Moneda>
): PendienteDetalleUi {
    val items = state.cuotasPendientesDetalle.map { detalle ->
        PendienteDetalleItem(
            cliente = detalle.cliente,
            idPrestamo = detalle.idPrestamo,
            numeroCuota = detalle.numeroCuota,
            fechaVencimiento = detalle.fechaVencimiento,
            saldoPendiente = detalle.saldoPendiente,
            moneda = detalle.moneda
        )
    }.sortedBy { it.fechaVencimiento }

    val totals = items
        .groupBy { it.moneda }
        .mapValues { (_, values) -> values.sumOf { it.saldoPendiente } }

    return PendienteDetalleUi(
        title = "Saldo pendiente",
        totalCuotas = items.size,
        totalsByCurrency = totals,
        visibleCurrencies = visibleCurrencies,
        items = items
    )
}

@Composable
private fun DashboardMoneyCard(
    title: String,
    totals: Map<Moneda, Double>,
    visibleCurrencies: List<Moneda>,
    modifier: Modifier = Modifier,
    highlightValue: Boolean = false,
    valueColor: Color? = null,
    onClick: () -> Unit
) {
    val ordered = visibleCurrencies.ifEmpty { listOf(Moneda.SOLES) }
    Card(
        modifier = modifier.clickable(onClick = onClick),
        colors = androidx.compose.material3.CardDefaults.cardColors(
            containerColor = MaterialTheme.colorScheme.secondaryContainer,
            contentColor = MaterialTheme.colorScheme.onSecondaryContainer
        )
    ) {
        Column(
            modifier = Modifier.padding(12.dp),
            verticalArrangement = Arrangement.spacedBy(8.dp)
        ) {
            Text(title, style = MaterialTheme.typography.labelLarge)
            Row(
                modifier = Modifier.fillMaxWidth(),
                horizontalArrangement = Arrangement.spacedBy(8.dp)
            ) {
                ordered.forEach { moneda ->
                    androidx.compose.material3.Surface(
                        modifier = Modifier.weight(1f),
                        shape = RoundedCornerShape(10.dp),
                        color = MaterialTheme.colorScheme.surface.copy(alpha = 0.42f)
                    ) {
                        Column(
                            modifier = Modifier.padding(horizontal = 8.dp, vertical = 8.dp),
                            horizontalAlignment = Alignment.CenterHorizontally,
                            verticalArrangement = Arrangement.spacedBy(2.dp)
                        ) {
                            Text(
                                text = (totals[moneda] ?: 0.0).toMoney(moneda),
                                style = MaterialTheme.typography.titleMedium,
                                color = valueColor ?: if (highlightValue) {
                                    MaterialTheme.colorScheme.tertiary
                                } else {
                                    MaterialTheme.colorScheme.onSecondaryContainer
                                },
                                maxLines = 1
                            )
                            Text(
                                text = moneda.displayName,
                                style = MaterialTheme.typography.labelSmall,
                                color = MaterialTheme.colorScheme.onSecondaryContainer.copy(alpha = 0.72f),
                                maxLines = 1,
                                overflow = TextOverflow.Ellipsis
                            )
                        }
                    }
                }
            }
        }
    }
}

private fun resolveVisibleCurrencies(mainCode: String?, secondaryCode: String?): List<Moneda> {
    val mapped = listOfNotNull(mainCode.toMonedaOrNull(), secondaryCode.toMonedaOrNull())
        .distinct()
    return if (mapped.isEmpty()) listOf(Moneda.SOLES) else mapped
}

private fun String?.toMonedaOrNull(): Moneda? = Moneda.fromCode(this)
private fun compartirTextoDetalle(context: android.content.Context, detalle: DashboardDetalleInfo) {
    val sendIntent = Intent(Intent.ACTION_SEND).apply {
        type = "text/plain"
        putExtra(Intent.EXTRA_SUBJECT, "Detalle dashboard: ${detalle.title}")
        putExtra(Intent.EXTRA_TEXT, "${detalle.title}\n${detalle.message}")
    }
    context.startActivity(Intent.createChooser(sendIntent, "Compartir detalle"))
}

private fun compartirTextoPlano(context: android.content.Context, titulo: String, detalle: String) {
    val sendIntent = Intent(Intent.ACTION_SEND).apply {
        type = "text/plain"
        putExtra(Intent.EXTRA_SUBJECT, titulo)
        putExtra(Intent.EXTRA_TEXT, detalle)
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
