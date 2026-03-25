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
import androidx.compose.foundation.layout.size
import androidx.compose.foundation.layout.width
import androidx.compose.foundation.rememberScrollState
import androidx.compose.foundation.lazy.LazyColumn
import androidx.compose.foundation.lazy.items
import androidx.compose.foundation.shape.RoundedCornerShape
import androidx.compose.foundation.verticalScroll
import androidx.compose.material3.Card
import androidx.compose.material3.CardDefaults
import androidx.compose.material3.Button
import androidx.compose.material3.Icon
import androidx.compose.material3.MaterialTheme
import androidx.compose.material3.OutlinedTextField
import androidx.compose.material3.Text
import androidx.compose.foundation.text.KeyboardOptions
import androidx.compose.material.icons.Icons
import androidx.compose.material.icons.outlined.AttachMoney
import androidx.compose.material.icons.outlined.Payments
import androidx.compose.material.icons.outlined.Schedule
import androidx.compose.material.icons.outlined.TrendingUp
import androidx.compose.material.icons.outlined.Warning
import androidx.compose.runtime.Composable
import androidx.compose.runtime.getValue
import androidx.compose.runtime.mutableStateOf
import androidx.compose.runtime.remember
import androidx.compose.runtime.setValue
import androidx.compose.ui.platform.LocalContext
import androidx.compose.ui.Alignment
import androidx.compose.ui.Modifier
import androidx.compose.ui.graphics.Color
import androidx.compose.ui.graphics.vector.ImageVector
import androidx.compose.ui.text.font.FontWeight
import androidx.compose.ui.text.input.KeyboardType
import androidx.compose.ui.text.style.TextOverflow
import androidx.compose.ui.text.style.TextAlign
import com.prestamos.app.ui.theme.AccentGold
import com.prestamos.app.ui.theme.PrimaryGreen
import com.prestamos.app.ui.theme.SecondaryGreen
import androidx.compose.ui.unit.dp
import androidx.compose.ui.unit.sp
import androidx.core.content.FileProvider
import androidx.lifecycle.compose.collectAsStateWithLifecycle
import com.prestamos.app.data.config.InitialSetupPreferences
import com.prestamos.app.data.license.LicenseType
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
    onRefreshLicenseStatus: () -> Unit,
    onGoToActivation: () -> Unit
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
                            "\uD83D\uDC64 $businessName",
                            style = MaterialTheme.typography.titleMedium,
                            color = MaterialTheme.colorScheme.primary
                        )
                    }
                    Text("\uD83D\uDCC5 ${System.currentTimeMillis().toDateString()}")
                    Spacer(modifier = Modifier.height(8.dp))
                    DashboardLicenseStatusCard(
                        activationUiState = activationUiState,
                        onGoToActivation = onGoToActivation
                    )
                }
            }
        }

        item {
            val capitalPrestadoActivo2 = state.prestamosActivosDetalle.sumByCurrency { it.montoPrestado }
            val prestadoActivoConInteres = state.prestamosActivosDetalle.sumByCurrency { it.montoTotalConInteres }
            Text(
                text = "\uD83D\uDCDC Historial de prestamos",
                color = MaterialTheme.colorScheme.primary,
                modifier = Modifier.clickable { detalleSeleccionado = DashboardDetalle.HISTORIAL }
            )
            Spacer(modifier = Modifier.height(8.dp))
            DashboardMoneyCard(
                title = "Capital prestado (activo)",
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
                title = "Total prestado activo (con intereses)",
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
                    title = "Pendiente",
                    totals = state.cuotasPendientesDetalle.sumByCurrency { it.saldoPendiente },
                    visibleCurrencies = visibleCurrencies,
                    modifier = Modifier.weight(1f),
                    stacked = true,
                    compact = true
                ) {
                    detalleSeleccionado = DashboardDetalle.PENDIENTE
                }
                DashboardMoneyCard(
                    title = "Cobrado hoy",
                    totals = state.pagosHoyDetalle.sumByCurrency { it.montoAbono },
                    visibleCurrencies = visibleCurrencies,
                    modifier = Modifier.weight(1f),
                    highlightValue = true,
                    valueColor = MaterialTheme.colorScheme.onSecondaryContainer,
                    stacked = true,
                    compact = true
                ) {
                    detalleSeleccionado = DashboardDetalle.COBRADO_HOY
                }
            }
            Spacer(modifier = Modifier.height(8.dp))
            DashboardCard(
                title = "Cuotas vencidas",
                value = state.cuotasVencidas.toString(),
                modifier = Modifier.fillMaxWidth(),
                centerContent = true
            ) {
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
                totals = state.gananciasPrestamosPagados.sumByCurrency { it.ganancia + it.moraCobrada },
                visibleCurrencies = visibleCurrencies,
                modifier = Modifier.fillMaxWidth()
            ) {
                detalleSeleccionado = DashboardDetalle.GANANCIAS
            }

            Spacer(Modifier.height(8.dp))
            if (state.gananciasPrestamosPagados.isEmpty()) {
                Text("No hay prestamos pagados aun")
            } else {
                val maxGanancia = state.gananciasPrestamosPagados.maxOf { it.ganancia + it.moraCobrada }.coerceAtLeast(1.0)
                Column(verticalArrangement = Arrangement.spacedBy(8.dp)) {
                    state.gananciasPrestamosPagados.take(8).forEach { item ->
                        val totalGanado = item.ganancia + item.moraCobrada
                        val ratio = (totalGanado / maxGanancia).toFloat().coerceIn(0.1f, 1f)
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
                            "${item.cliente}: total ganado ${totalGanado.toMoney(item.moneda)}",
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
        } else if (it == DashboardDetalle.HISTORIAL) {
            val historialDetalle = it.toHistorialDetalleUi(state, visibleCurrencies)
            HistorialDetalleDialog(
                detalle = historialDetalle,
                onClose = { detalleSeleccionado = null },
                onShareText = {
                    compartirTextoPlano(
                        context = context,
                        titulo = historialDetalle.title,
                        detalle = historialDetalle.toShareText()
                    )
                },
                onSharePdf = {
                    runCatching {
                        createDashboardDetallePdf(context, historialDetalle.title, historialDetalle.toShareText())
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
        } else if (it == DashboardDetalle.COBRADO_HOY) {
            val cobradoDetalle = it.toCobradoHoyDetalleUi(state, visibleCurrencies)
            CobradoHoyDetalleDialog(
                detalle = cobradoDetalle,
                onClose = { detalleSeleccionado = null },
                onShareText = {
                    compartirTextoPlano(
                        context = context,
                        titulo = cobradoDetalle.title,
                        detalle = cobradoDetalle.toShareText()
                    )
                },
                onSharePdf = {
                    runCatching {
                        createDashboardDetallePdf(context, cobradoDetalle.title, cobradoDetalle.toShareText())
                    }.onSuccess { file ->
                        compartirArchivoDetalle(context, file, "application/pdf")
                    }.onFailure {
                        Toast.makeText(context, "No se pudo exportar PDF", Toast.LENGTH_SHORT).show()
                    }
                }
            )
        } else if (it == DashboardDetalle.GANANCIAS) {
            val gananciasDetalle = it.toGananciasDetalleUi(state, visibleCurrencies)
            GananciasDetalleDialog(
                detalle = gananciasDetalle,
                onClose = { detalleSeleccionado = null },
                onShareText = {
                    compartirTextoPlano(
                        context = context,
                        titulo = gananciasDetalle.title,
                        detalle = gananciasDetalle.toShareText()
                    )
                },
                onSharePdf = {
                    runCatching {
                        createDashboardDetallePdf(context, gananciasDetalle.title, gananciasDetalle.toShareText())
                    }.onSuccess { file ->
                        compartirArchivoDetalle(context, file, "application/pdf")
                    }.onFailure {
                        Toast.makeText(context, "No se pudo exportar PDF", Toast.LENGTH_SHORT).show()
                    }
                }
            )
        } else if (it == DashboardDetalle.VENCIDAS) {
            val vencidasDetalle = it.toVencidasDetalleUi(state)
            VencidasDetalleDialog(
                detalle = vencidasDetalle,
                onClose = { detalleSeleccionado = null },
                onApplyMora = { idCuota, montoMora, onDone ->
                    viewModel.aplicarMoraCuotaVencida(
                        idCuota = idCuota,
                        montoMora = montoMora,
                        onSuccess = {
                            Toast.makeText(context, "Mora aplicada correctamente", Toast.LENGTH_SHORT).show()
                            onDone()
                        },
                        onError = {
                            Toast.makeText(context, it, Toast.LENGTH_SHORT).show()
                        }
                    )
                },
                onShareText = {
                    compartirTextoPlano(
                        context = context,
                        titulo = vencidasDetalle.title,
                        detalle = vencidasDetalle.toShareText()
                    )
                },
                onSharePdf = {
                    runCatching {
                        createDashboardDetallePdf(context, vencidasDetalle.title, vencidasDetalle.toShareText())
                    }.onSuccess { file ->
                        compartirArchivoDetalle(context, file, "application/pdf")
                    }.onFailure {
                        Toast.makeText(context, "No se pudo exportar PDF", Toast.LENGTH_SHORT).show()
                    }
                }
            )
        } else if (it == DashboardDetalle.COBRADO_ACTIVO) {
            val cobradoActivoDetalle = it.toCobradoActivoDetalleUi(state, visibleCurrencies)
            CobradoActivoDetalleDialog(
                detalle = cobradoActivoDetalle,
                onClose = { detalleSeleccionado = null },
                onShareText = {
                    compartirTextoPlano(
                        context = context,
                        titulo = cobradoActivoDetalle.title,
                        detalle = cobradoActivoDetalle.toShareText()
                    )
                },
                onSharePdf = {
                    runCatching {
                        createDashboardDetallePdf(context, cobradoActivoDetalle.title, cobradoActivoDetalle.toShareText())
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
private fun DashboardLicenseStatusCard(
    activationUiState: ActivationUiState,
    onGoToActivation: () -> Unit
) {
    val statusUi = activationUiState.status.toDashboardLicenseUi()
    Card(
        modifier = Modifier.fillMaxWidth(),
        shape = RoundedCornerShape(12.dp),
        colors = CardDefaults.cardColors(containerColor = statusUi.containerColor)
    ) {
        Column(
            modifier = Modifier
                .fillMaxWidth()
                .padding(10.dp),
            verticalArrangement = Arrangement.spacedBy(4.dp)
        ) {
            Text(
                text = statusUi.title,
                style = MaterialTheme.typography.labelLarge,
                color = statusUi.titleColor
            )
            Text(
                text = statusUi.subtitle,
                style = MaterialTheme.typography.bodySmall,
                color = statusUi.subtitleColor
            )
            if (statusUi.showActivateButton) {
                Button(
                    onClick = onGoToActivation,
                    shape = RoundedCornerShape(10.dp),
                    modifier = Modifier.fillMaxWidth()
                ) {
                    Text("Activar version Pro")
                }
            }
        }
    }
}

private data class DashboardLicenseUi(
    val title: String,
    val subtitle: String,
    val containerColor: Color,
    val showActivateButton: Boolean,
    val titleColor: Color,
    val subtitleColor: Color
)

private fun com.prestamos.app.data.license.LicenseStatus.toDashboardLicenseUi(): DashboardLicenseUi {
    if (licenseType == LicenseType.TRIAL && !trialExpired) {
        return DashboardLicenseUi(
            title = "Versión Lite",
            subtitle = "Activa version Pro para desbloquear todas las funciones",
            containerColor = Color(0xFFFFF3CC),
            showActivateButton = true,
            titleColor = Color(0xFF7A4F01),
            subtitleColor = Color(0xFF8A5A00)
        )
    }

    if (isValid && isActivated) {
        val plan = when (licenseType) {
            LicenseType.MENSUAL -> "Mensual"
            LicenseType.ANUAL -> "Anual"
            LicenseType.FULL -> "Full"
            LicenseType.TRIAL -> "Prueba"
        }
        val vigencia = if (licenseType == LicenseType.FULL) {
            "Sin vencimiento"
        } else {
            expirationDate?.toDateString() ?: "No disponible"
        }
        return DashboardLicenseUi(
            title = "Licencia activa",
            subtitle = "Plan $plan - Valida hasta $vigencia",
            containerColor = Color(0xFFDDF3E1),
            showActivateButton = false,
            titleColor = Color(0xFF1B5E20),
            subtitleColor = Color(0xFF2E7D32)
        )
    }

    return DashboardLicenseUi(
        title = "Licencia expirada",
        subtitle = "Tu acceso completo termino",
        containerColor = Color(0xFFFFDCDC),
        showActivateButton = true,
        titleColor = Color(0xFF8E1F1F),
        subtitleColor = Color(0xFF9D2B2B)
    )
}

@Composable
private fun DashboardCard(
    title: String,
    value: String,
    modifier: Modifier = Modifier,
    highlightValue: Boolean = false,
    valueColor: Color? = null,
    centerContent: Boolean = false,
    onClick: () -> Unit
) {
    Card(
        modifier = modifier.clickable(onClick = onClick),
        colors = androidx.compose.material3.CardDefaults.cardColors(
            containerColor = MaterialTheme.colorScheme.secondaryContainer,
            contentColor = MaterialTheme.colorScheme.onSecondaryContainer
        )
    ) {
        Column(
            modifier = Modifier.padding(12.dp),
            horizontalAlignment = if (centerContent) Alignment.CenterHorizontally else Alignment.Start
        ) {
            DashboardCardTitle(
                title = title,
                centerContent = centerContent,
                compact = false
            )
            Text(
                value,
                style = MaterialTheme.typography.titleLarge,
                color = valueColor ?: if (highlightValue) {
                    MaterialTheme.colorScheme.tertiary
                } else {
                    MaterialTheme.colorScheme.onSecondaryContainer
                },
                textAlign = if (centerContent) TextAlign.Center else TextAlign.Start,
                modifier = if (centerContent) Modifier.fillMaxWidth() else Modifier
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
            Text("$iconPin Cuota ${cuota.numeroCuota} \u2022 \uD83D\uDCC4 Prestamo #${cuota.idPrestamo}", style = MaterialTheme.typography.bodyMedium, color = Color(0xFF1F1F1F))
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
    val iconTipoCobro = "\uD83D\uDCB8"

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
            Text("$iconPin Cuota ${pago.numeroCuota} \u2022 \uD83D\uDCC4 Prestamo #${pago.idPrestamo}", style = MaterialTheme.typography.bodyMedium)
            Text("$iconCal ${pago.fechaPago.toDateString()}", style = MaterialTheme.typography.bodyMedium)
            Text("$iconTipoCobro ${pago.tipoCobro}", style = MaterialTheme.typography.bodyMedium)
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

private data class HistorialDetalleItem(
    val cliente: String,
    val idPrestamo: Long,
    val fechaRegistro: Long,
    val montoCobrado: Double,
    val montoGanado: Double,
    val moneda: Moneda
)

private data class HistorialDetalleUi(
    val title: String,
    val totalPrestamosPagados: Int,
    val totalCobradoByCurrency: Map<Moneda, Double>,
    val totalGanadoByCurrency: Map<Moneda, Double>,
    val visibleCurrencies: List<Moneda>,
    val items: List<HistorialDetalleItem>
)

private data class PendienteDetalleItem(
    val cliente: String,
    val idPrestamo: Long,
    val numeroCuota: Int,
    val fechaVencimiento: Long,
    val saldoPendiente: Double,
    val moraPendiente: Double,
    val moneda: Moneda
)

private data class PendienteDetalleUi(
    val title: String,
    val totalCuotas: Int,
    val totalsByCurrency: Map<Moneda, Double>,
    val visibleCurrencies: List<Moneda>,
    val items: List<PendienteDetalleItem>
)

private data class CobradoHoyDetalleItem(
    val cliente: String,
    val idPrestamo: Long,
    val numeroCuota: Int,
    val fechaPago: Long,
    val montoAbono: Double,
    val moneda: Moneda,
    val tipoCobro: String
)

private data class CobradoHoyDetalleUi(
    val title: String,
    val totalPagos: Int,
    val totalsByCurrency: Map<Moneda, Double>,
    val visibleCurrencies: List<Moneda>,
    val items: List<CobradoHoyDetalleItem>
)

private data class GananciaDetalleItem(
    val cliente: String,
    val idPrestamo: Long,
    val numeroCuota: Int,
    val fechaPago: Long,
    val ganancia: Double,
    val moraCobrada: Double,
    val moneda: Moneda
)

private data class GananciasDetalleUi(
    val title: String,
    val totalPrestamosPagados: Int,
    val totalsByCurrency: Map<Moneda, Double>,
    val visibleCurrencies: List<Moneda>,
    val items: List<GananciaDetalleItem>
)

private data class VencidaDetalleItem(
    val cliente: String,
    val idCuota: Long,
    val idPrestamo: Long,
    val numeroCuota: Int,
    val fechaVencimiento: Long,
    val saldoPendiente: Double,
    val moraPendiente: Double,
    val moneda: Moneda
)

private data class VencidasDetalleUi(
    val title: String,
    val totalCuotasVencidas: Int,
    val items: List<VencidaDetalleItem>
)

private data class CobradoActivoDetalleItem(
    val cliente: String,
    val idPrestamo: Long,
    val numeroCuota: Int,
    val fechaPago: Long,
    val montoAbono: Double,
    val moneda: Moneda,
    val tipoCobro: String
)

private data class CobradoActivoDetalleUi(
    val title: String,
    val totalPrestamosActivos: Int,
    val totalsByCurrency: Map<Moneda, Double>,
    val visibleCurrencies: List<Moneda>,
    val items: List<CobradoActivoDetalleItem>
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

private fun HistorialDetalleUi.toShareText(): String = buildString {
    appendLine(title)
    appendLine("Prestamos pagados: $totalPrestamosPagados")
    appendLine("Total cobrado:")
    visibleCurrencies.forEach { moneda ->
        appendLine("- ${moneda.displayName}: ${(totalCobradoByCurrency[moneda] ?: 0.0).toMoney(moneda)}")
    }
    appendLine("Total ganado:")
    visibleCurrencies.forEach { moneda ->
        appendLine("- ${moneda.displayName}: ${(totalGanadoByCurrency[moneda] ?: 0.0).toMoney(moneda)}")
    }
    appendLine()
    appendLine("Detalle")
    if (items.isEmpty()) {
        appendLine("No hay prestamos pagados.")
    } else {
        items.forEach { item ->
            appendLine("${item.cliente} | Prestamo #${item.idPrestamo} | Fecha ${item.fechaRegistro.toDateString()} | Cobrado ${item.montoCobrado.toMoney(item.moneda)} ${item.moneda.displayName} | Ganado ${item.montoGanado.toMoney(item.moneda)} ${item.moneda.displayName}")
        }
    }
}

private fun CobradoActivoDetalleUi.toShareText(): String = buildString {
    appendLine(title)
    appendLine("Prestamos activos: $totalPrestamosActivos")
    visibleCurrencies.forEach { moneda ->
        appendLine("Total ${moneda.displayName}: ${(totalsByCurrency[moneda] ?: 0.0).toMoney(moneda)}")
    }
    appendLine()
    appendLine("Detalle")
    if (items.isEmpty()) {
        appendLine("No hay cobros registrados en prestamos activos.")
    } else {
        items.forEach { item ->
            appendLine("${item.cliente} | Prestamo #${item.idPrestamo} | Cuota ${item.numeroCuota} | Fecha ${item.fechaPago.toDateString()} | Tipo cobro ${item.tipoCobro} | Cobro ${item.montoAbono.toMoney(item.moneda)} ${item.moneda.displayName}")
        }
    }
}

private fun VencidasDetalleUi.toShareText(): String = buildString {
    appendLine(title)
    appendLine("Cuotas vencidas: $totalCuotasVencidas")
    appendLine()
    appendLine("Detalle")
    if (items.isEmpty()) {
        appendLine("No hay cuotas vencidas.")
    } else {
        items.forEach { item ->
            if (item.moraPendiente > 0.0) {
                val saldoFinal = item.saldoPendiente + item.moraPendiente
                appendLine("${item.cliente} | Prestamo #${item.idPrestamo} | Cuota ${item.numeroCuota} | Vence ${item.fechaVencimiento.toDateString()} | Saldo ${item.saldoPendiente.toMoney(item.moneda)} | Mora ${item.moraPendiente.toMoney(item.moneda)} | Saldo final ${saldoFinal.toMoney(item.moneda)} ${item.moneda.displayName}")
            } else {
                appendLine("${item.cliente} | Prestamo #${item.idPrestamo} | Cuota ${item.numeroCuota} | Vence ${item.fechaVencimiento.toDateString()} | Saldo ${item.saldoPendiente.toMoney(item.moneda)} ${item.moneda.displayName}")
            }
        }
    }
}

private fun GananciasDetalleUi.toShareText(): String = buildString {
    appendLine(title)
    appendLine("Prestamos pagados: $totalPrestamosPagados")
    visibleCurrencies.forEach { moneda ->
        appendLine("Total ganado ${moneda.displayName}: ${(totalsByCurrency[moneda] ?: 0.0).toMoney(moneda)}")
    }
    appendLine()
    appendLine("Detalle")
    if (items.isEmpty()) {
        appendLine("No hay prestamos pagados.")
    } else {
        items.forEach { item ->
            appendLine("${item.cliente} | Prestamo #${item.idPrestamo} | Cuota ${item.numeroCuota} | Fecha ${item.fechaPago.toDateString()} | Ganancia ${item.ganancia.toMoney(item.moneda)} | Mora ${item.moraCobrada.toMoney(item.moneda)} ${item.moneda.displayName}")
        }
    }
}

private fun CobradoHoyDetalleUi.toShareText(): String = buildString {
    appendLine(title)
    appendLine("Pagos de hoy: $totalPagos")
    visibleCurrencies.forEach { moneda ->
        appendLine("Total cobrado ${moneda.displayName}: ${(totalsByCurrency[moneda] ?: 0.0).toMoney(moneda)}")
    }
    appendLine()
    appendLine("Detalle")
    if (items.isEmpty()) {
        appendLine("No hay pagos registrados hoy.")
    } else {
        items.forEach { item ->
            appendLine("${item.cliente} | Prestamo #${item.idPrestamo} | Cuota ${item.numeroCuota} | Fecha ${item.fechaPago.toDateString()} | Tipo cobro ${item.tipoCobro} | Cobro ${item.montoAbono.toMoney(item.moneda)} ${item.moneda.displayName}")
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
            if (item.moraPendiente > 0.0) {
                val saldoFinal = item.saldoPendiente + item.moraPendiente
                appendLine("${item.cliente} | Prestamo #${item.idPrestamo} | Cuota ${item.numeroCuota} | Vence ${item.fechaVencimiento.toDateString()} | Saldo ${item.saldoPendiente.toMoney(item.moneda)} | Mora ${item.moraPendiente.toMoney(item.moneda)} | Saldo final ${saldoFinal.toMoney(item.moneda)} ${item.moneda.displayName}")
            } else {
                appendLine("${item.cliente} | Prestamo #${item.idPrestamo} | Cuota ${item.numeroCuota} | Vence ${item.fechaVencimiento.toDateString()} | Saldo ${item.saldoPendiente.toMoney(item.moneda)} ${item.moneda.displayName}")
            }
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
                            Text("\uD83D\uDCB0 Total ${moneda.displayName}", style = MaterialTheme.typography.bodyMedium)
                            Text((detalle.totalsByCurrency[moneda] ?: 0.0).toMoney(moneda), style = MaterialTheme.typography.bodyMedium)
                        }
                    }
                }
                Text("\uD83D\uDCC4 Detalle de prestamos", style = MaterialTheme.typography.labelLarge)
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
                                    Text("\uD83D\uDC64 ${item.cliente}", style = MaterialTheme.typography.bodyMedium)
                                    Text("\uD83D\uDCB0 ${item.monto.toMoney(item.moneda)} ${item.moneda.displayName}", style = MaterialTheme.typography.bodySmall)
                                    Text("\uD83D\uDCC4 Prestamo #${item.idPrestamo}", style = MaterialTheme.typography.bodySmall)
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
private fun HistorialDetalleDialog(
    detalle: HistorialDetalleUi,
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
                        text = "\uD83D\uDD22 Prestamos pagados: ${detalle.totalPrestamosPagados}",
                        modifier = Modifier.padding(horizontal = 8.dp, vertical = 6.dp),
                        style = MaterialTheme.typography.bodyMedium
                    )
                }

                Text("Total cobrado", style = MaterialTheme.typography.labelLarge)
                detalle.visibleCurrencies.forEach { moneda ->
                    Card(
                        modifier = Modifier.fillMaxWidth(),
                        colors = CardDefaults.cardColors(
                            containerColor = MaterialTheme.colorScheme.surfaceVariant.copy(alpha = 0.55f)
                        )
                    ) {
                        Text(
                            text = "\uD83D\uDCB0 Total ${moneda.displayName}: ${(detalle.totalCobradoByCurrency[moneda] ?: 0.0).toMoney(moneda)}",
                            modifier = Modifier.padding(horizontal = 8.dp, vertical = 6.dp),
                            style = MaterialTheme.typography.bodyMedium
                        )
                    }
                }

                Text("Total ganado", style = MaterialTheme.typography.labelLarge)
                detalle.visibleCurrencies.forEach { moneda ->
                    Card(
                        modifier = Modifier.fillMaxWidth(),
                        colors = CardDefaults.cardColors(
                            containerColor = MaterialTheme.colorScheme.surfaceVariant.copy(alpha = 0.55f)
                        )
                    ) {
                        Text(
                            text = "\uD83D\uDCB0 Total ${moneda.displayName}: ${(detalle.totalGanadoByCurrency[moneda] ?: 0.0).toMoney(moneda)}",
                            modifier = Modifier.padding(horizontal = 8.dp, vertical = 6.dp),
                            style = MaterialTheme.typography.bodyMedium
                        )
                    }
                }

                Text("\uD83D\uDCCB Detalle", style = MaterialTheme.typography.labelLarge)
                if (detalle.items.isEmpty()) {
                    Text("No hay prestamos pagados.")
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
                                    Text("\uD83D\uDCC4 Prestamo #${item.idPrestamo}", style = MaterialTheme.typography.bodySmall)
                                    Text("\uD83D\uDCC5 Fecha ${item.fechaRegistro.toDateString()}", style = MaterialTheme.typography.bodySmall)
                                    Text("\uD83D\uDCB0 Cobrado ${item.montoCobrado.toMoney(item.moneda)} ${item.moneda.displayName}", style = MaterialTheme.typography.bodySmall)
                                    Text("\uD83D\uDCB0 Ganado ${item.montoGanado.toMoney(item.moneda)} ${item.moneda.displayName}", style = MaterialTheme.typography.bodySmall)
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
                        text = "\uD83D\uDCCC Cuotas pendientes: ${detalle.totalCuotas}",
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
                                    Text("\uD83D\uDCC4 Prestamo #${item.idPrestamo} \u00B7 \uD83D\uDCCC Cuota ${item.numeroCuota}", style = MaterialTheme.typography.bodySmall)
                                    Text("\uD83D\uDCC5 Vence ${item.fechaVencimiento.toDateString()}", style = MaterialTheme.typography.bodySmall)
                                    Text("\uD83D\uDCB0 Saldo ${item.saldoPendiente.toMoney(item.moneda)} ${item.moneda.displayName}", style = MaterialTheme.typography.bodySmall)
                                    if (item.moraPendiente > 0.0) {
                                        val saldoFinal = item.saldoPendiente + item.moraPendiente
                                        Text("\uD83D\uDCB8 Mora ${item.moraPendiente.toMoney(item.moneda)} ${item.moneda.displayName}", style = MaterialTheme.typography.bodySmall)
                                        Text("\uD83D\uDCB5 Saldo final ${saldoFinal.toMoney(item.moneda)} ${item.moneda.displayName}", style = MaterialTheme.typography.bodySmall)
                                    }
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
private fun CobradoHoyDetalleDialog(
    detalle: CobradoHoyDetalleUi,
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
                        text = "\uD83D\uDD22 Pagos de hoy: ${detalle.totalPagos}",
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
                            text = "\uD83D\uDCB0 Total cobrado ${moneda.displayName}: ${(detalle.totalsByCurrency[moneda] ?: 0.0).toMoney(moneda)}",
                            modifier = Modifier.padding(horizontal = 8.dp, vertical = 6.dp),
                            style = MaterialTheme.typography.bodyMedium
                        )
                    }
                }
                Text("\uD83D\uDCCB Detalle", style = MaterialTheme.typography.labelLarge)
                if (detalle.items.isEmpty()) {
                    Text("No hay pagos registrados hoy.")
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
                                    Text("\uD83D\uDCC4 Prestamo #${item.idPrestamo} \u00B7 \uD83D\uDCCC Cuota ${item.numeroCuota}", style = MaterialTheme.typography.bodySmall)
                                    Text("\uD83D\uDCC5 Fecha ${item.fechaPago.toDateString()}", style = MaterialTheme.typography.bodySmall)
                                    Text("\uD83D\uDCB8 ${item.tipoCobro}", style = MaterialTheme.typography.bodySmall)
                                    Text("\uD83D\uDCB0 Cobro ${item.montoAbono.toMoney(item.moneda)} ${item.moneda.displayName}", style = MaterialTheme.typography.bodySmall)
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
private fun GananciasDetalleDialog(
    detalle: GananciasDetalleUi,
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
                        text = "\uD83D\uDD22 Prestamos pagados: ${detalle.totalPrestamosPagados}",
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
                            text = "\uD83D\uDCB0 Total ganado ${moneda.displayName}: ${(detalle.totalsByCurrency[moneda] ?: 0.0).toMoney(moneda)}",
                            modifier = Modifier.padding(horizontal = 8.dp, vertical = 6.dp),
                            style = MaterialTheme.typography.bodyMedium
                        )
                    }
                }
                Text("\uD83D\uDCCB Detalle", style = MaterialTheme.typography.labelLarge)
                if (detalle.items.isEmpty()) {
                    Text("No hay prestamos pagados.")
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
                                    Text("\uD83D\uDCC4 Prestamo #${item.idPrestamo} \u00B7 \uD83D\uDCCC Cuota ${item.numeroCuota}", style = MaterialTheme.typography.bodySmall)
                                    Text("\uD83D\uDCC5 Fecha ${item.fechaPago.toDateString()}", style = MaterialTheme.typography.bodySmall)
                                    Text("\uD83D\uDCB0 Ganancia ${item.ganancia.toMoney(item.moneda)} ${item.moneda.displayName}", style = MaterialTheme.typography.bodySmall)
                                    Text("\uD83D\uDCB8 Mora cobrada ${item.moraCobrada.toMoney(item.moneda)} ${item.moneda.displayName}", style = MaterialTheme.typography.bodySmall)
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
private fun VencidasDetalleDialog(
    detalle: VencidasDetalleUi,
    onClose: () -> Unit,
    onApplyMora: (idCuota: Long, montoMora: String, onDone: () -> Unit) -> Unit,
    onShareText: () -> Unit,
    onSharePdf: () -> Unit
) {
    var cuotaSeleccionadaMora by remember { mutableStateOf<VencidaDetalleItem?>(null) }
    var montoMoraInput by remember { mutableStateOf("") }

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
                        text = "\uD83D\uDCCC Cuotas vencidas: ${detalle.totalCuotasVencidas}",
                        modifier = Modifier.padding(horizontal = 8.dp, vertical = 6.dp),
                        style = MaterialTheme.typography.bodyMedium
                    )
                }
                Text("\uD83D\uDCCB Detalle", style = MaterialTheme.typography.labelLarge)
                if (detalle.items.isEmpty()) {
                    Text("No hay cuotas vencidas.")
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
                                    Text("\uD83D\uDCC4 Prestamo #${item.idPrestamo} \u00B7 \uD83D\uDCCC Cuota ${item.numeroCuota}", style = MaterialTheme.typography.bodySmall)
                                    val fechaVenc = Instant.ofEpochMilli(item.fechaVencimiento).atZone(ZoneId.systemDefault()).toLocalDate()
                                    val diasVencida = ChronoUnit.DAYS.between(fechaVenc, LocalDate.now()).coerceAtLeast(0L)
                                    Text("\uD83D\uDCC5 Vence ${item.fechaVencimiento.toDateString()}", style = MaterialTheme.typography.bodySmall)
                                    Text(
                                        text = if (diasVencida == 0L) "\u23F0 Vence hoy" else "\u23F0 Vencida hace $diasVencida ${if (diasVencida == 1L) "dia" else "dias"}",
                                        style = MaterialTheme.typography.bodySmall
                                    )
                                    if (item.moraPendiente > 0.0) {
                                        val saldoFinal = item.saldoPendiente + item.moraPendiente
                                        Text("\uD83D\uDCB0 Saldo: ${item.saldoPendiente.toMoney(item.moneda)} ${item.moneda.displayName}", style = MaterialTheme.typography.bodySmall)
                                        Text("\uD83D\uDCB8 Mora: ${item.moraPendiente.toMoney(item.moneda)} ${item.moneda.displayName}", style = MaterialTheme.typography.bodySmall)
                                        Text("\uD83D\uDCB5 Saldo final: ${saldoFinal.toMoney(item.moneda)} ${item.moneda.displayName}", style = MaterialTheme.typography.bodySmall)
                                    } else {
                                        Text("\uD83D\uDCB0 Saldo: ${item.saldoPendiente.toMoney(item.moneda)} ${item.moneda.displayName}", style = MaterialTheme.typography.bodySmall)
                                    }
                                    androidx.compose.material3.TextButton(
                                        onClick = {
                                            cuotaSeleccionadaMora = item
                                            montoMoraInput = ""
                                        },
                                        modifier = Modifier.align(Alignment.End)
                                    ) {
                                        Text("Aplicar mora")
                                    }
                                }
                            }
                        }
                    }
                }
            }
        }
    )

    if (cuotaSeleccionadaMora != null) {
        val cuota = cuotaSeleccionadaMora ?: return
        androidx.compose.material3.AlertDialog(
            onDismissRequest = { cuotaSeleccionadaMora = null },
            confirmButton = {
                androidx.compose.material3.TextButton(
                    onClick = {
                        onApplyMora(cuota.idCuota, montoMoraInput) {
                            cuotaSeleccionadaMora = null
                            montoMoraInput = ""
                        }
                    }
                ) {
                    Text("Aplicar")
                }
            },
            dismissButton = {
                androidx.compose.material3.TextButton(onClick = { cuotaSeleccionadaMora = null }) {
                    Text("Cancelar")
                }
            },
            title = { Text("Aplicar mora") },
            text = {
                val fechaVenc = Instant.ofEpochMilli(cuota.fechaVencimiento).atZone(ZoneId.systemDefault()).toLocalDate()
                val diasVencida = ChronoUnit.DAYS.between(fechaVenc, LocalDate.now()).coerceAtLeast(0L)
                Column(verticalArrangement = Arrangement.spacedBy(8.dp)) {
                    Text("\uD83D\uDCC4 Prestamo #${cuota.idPrestamo} \u00B7 \uD83D\uDCCC Cuota ${cuota.numeroCuota}")
                    Text("\uD83D\uDC64 Cliente: ${cuota.cliente}")
                    Text(
                        if (diasVencida == 0L) "\u23F0 Dias de vencimiento: vence hoy"
                        else "\u23F0 Dias de vencimiento: $diasVencida"
                    )
                    Text("Saldo actual: ${cuota.saldoPendiente.toMoney(cuota.moneda)}")
                    OutlinedTextField(
                        value = montoMoraInput,
                        onValueChange = { value ->
                            val clean = value.replace(',', '.')
                            val filtered = buildString {
                                var hasDot = false
                                clean.forEach { ch ->
                                    when {
                                        ch.isDigit() -> append(ch)
                                        ch == '.' && !hasDot -> {
                                            hasDot = true
                                            append(ch)
                                        }
                                    }
                                }
                            }
                            montoMoraInput = filtered
                        },
                        label = { Text("Monto mora") },
                        modifier = Modifier.fillMaxWidth(),
                        singleLine = true,
                        keyboardOptions = KeyboardOptions(keyboardType = KeyboardType.Decimal),
                        textStyle = MaterialTheme.typography.bodyLarge.copy(textAlign = TextAlign.End)
                    )
                }
            }
        )
    }
}

@Composable
private fun CobradoActivoDetalleDialog(
    detalle: CobradoActivoDetalleUi,
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
                        text = "\uD83D\uDD22 Prestamos activos: ${detalle.totalPrestamosActivos}",
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
                            text = "\uD83D\uDCB0 Total ${moneda.displayName}: ${(detalle.totalsByCurrency[moneda] ?: 0.0).toMoney(moneda)}",
                            modifier = Modifier.padding(horizontal = 8.dp, vertical = 6.dp),
                            style = MaterialTheme.typography.bodyMedium
                        )
                    }
                }
                Text("\uD83D\uDCCB Detalle", style = MaterialTheme.typography.labelLarge)
                if (detalle.items.isEmpty()) {
                    Text("No hay cobros registrados en prestamos activos.")
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
                                    Text("\uD83D\uDCC4 Prestamo #${item.idPrestamo} \u00B7 \uD83D\uDCCC Cuota ${item.numeroCuota}", style = MaterialTheme.typography.bodySmall)
                                    Text("\uD83D\uDCC5 Fecha ${item.fechaPago.toDateString()}", style = MaterialTheme.typography.bodySmall)
                                    Text("\uD83D\uDCB8 ${item.tipoCobro}", style = MaterialTheme.typography.bodySmall)
                                    Text("\uD83D\uDCB0 Cobro ${item.montoAbono.toMoney(item.moneda)} ${item.moneda.displayName}", style = MaterialTheme.typography.bodySmall)
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
                "- ${it.cliente} | Prestamo #${it.idPrestamo} | Cuota ${it.numeroCuota} | ${it.fechaPago.toDateString()} | Tipo cobro ${it.tipoCobro} | Abono ${it.montoAbono.toMoney(it.moneda)}"
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
            val gananciaPorMoneda = state.gananciasPrestamosPagados.sumByCurrency { it.ganancia + it.moraCobrada }
            val top = state.gananciasPrestamosPagados.take(12).joinToString("\n") {
                "- ${it.cliente} | Prestamo #${it.idPrestamo} | Prestado ${it.montoPrestado.toMoney(it.moneda)} | Cobrado ${it.montoCobrado.toMoney(it.moneda)} | Ganancia ${it.ganancia.toMoney(it.moneda)} | Mora ${it.moraCobrada.toMoney(it.moneda)}"
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

private fun DashboardDetalle.toHistorialDetalleUi(
    state: com.prestamos.app.ui.model.DashboardResumen,
    visibleCurrencies: List<Moneda>
): HistorialDetalleUi {
    val items = state.prestamosCapitalDetalle
        .filter { it.cuotasPendientes == 0 }
        .map { detalle ->
            HistorialDetalleItem(
                cliente = detalle.cliente,
                idPrestamo = detalle.idPrestamo,
                fechaRegistro = detalle.fechaRegistro,
                montoCobrado = detalle.montoCobrado,
                montoGanado = (detalle.montoCobrado - detalle.montoPrestado).coerceAtLeast(0.0),
                moneda = detalle.moneda
            )
        }
        .sortedByDescending { it.fechaRegistro }

    val totalCobrado = items
        .groupBy { it.moneda }
        .mapValues { (_, values) -> values.sumOf { it.montoCobrado } }

    val totalGanado = items
        .groupBy { it.moneda }
        .mapValues { (_, values) -> values.sumOf { it.montoGanado } }

    return HistorialDetalleUi(
        title = "Historial de prestamos",
        totalPrestamosPagados = items.size,
        totalCobradoByCurrency = totalCobrado,
        totalGanadoByCurrency = totalGanado,
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
            moraPendiente = detalle.moraPendiente,
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

private fun DashboardDetalle.toCobradoHoyDetalleUi(
    state: com.prestamos.app.ui.model.DashboardResumen,
    visibleCurrencies: List<Moneda>
): CobradoHoyDetalleUi {
    val items = state.pagosHoyDetalle.map { detalle ->
        CobradoHoyDetalleItem(
            cliente = detalle.cliente,
            idPrestamo = detalle.idPrestamo,
            numeroCuota = detalle.numeroCuota,
            fechaPago = detalle.fechaPago,
            montoAbono = detalle.montoAbono,
            moneda = detalle.moneda,
            tipoCobro = detalle.tipoCobro
        )
    }.sortedByDescending { it.fechaPago }

    val totals = items
        .groupBy { it.moneda }
        .mapValues { (_, values) -> values.sumOf { it.montoAbono } }

    return CobradoHoyDetalleUi(
        title = "Cobrado hoy",
        totalPagos = items.size,
        totalsByCurrency = totals,
        visibleCurrencies = visibleCurrencies,
        items = items
    )
}

private fun DashboardDetalle.toGananciasDetalleUi(
    state: com.prestamos.app.ui.model.DashboardResumen,
    visibleCurrencies: List<Moneda>
): GananciasDetalleUi {
    val items = state.gananciasPrestamosPagados.map { detalle ->
        GananciaDetalleItem(
            cliente = detalle.cliente,
            idPrestamo = detalle.idPrestamo,
            numeroCuota = detalle.numeroCuota,
            fechaPago = detalle.fechaPago,
            ganancia = detalle.ganancia,
            moraCobrada = detalle.moraCobrada,
            moneda = detalle.moneda
        )
    }.sortedByDescending { it.fechaPago }

    val totals = items
        .groupBy { it.moneda }
        .mapValues { (_, values) -> values.sumOf { it.ganancia + it.moraCobrada } }

    return GananciasDetalleUi(
        title = "Ganancias por prestamos pagados",
        totalPrestamosPagados = items.size,
        totalsByCurrency = totals,
        visibleCurrencies = visibleCurrencies,
        items = items
    )
}

private fun DashboardDetalle.toVencidasDetalleUi(
    state: com.prestamos.app.ui.model.DashboardResumen
): VencidasDetalleUi {
    val items = state.cuotasVencidasDetalle.map { detalle ->
        VencidaDetalleItem(
            cliente = detalle.cliente,
            idCuota = detalle.idCuota,
            idPrestamo = detalle.idPrestamo,
            numeroCuota = detalle.numeroCuota,
            fechaVencimiento = detalle.fechaVencimiento,
            saldoPendiente = detalle.saldoPendiente,
            moraPendiente = detalle.moraPendiente,
            moneda = detalle.moneda
        )
    }.sortedBy { it.fechaVencimiento }

    return VencidasDetalleUi(
        title = "Cuotas vencidas",
        totalCuotasVencidas = items.size,
        items = items
    )
}

private fun DashboardDetalle.toCobradoActivoDetalleUi(
    state: com.prestamos.app.ui.model.DashboardResumen,
    visibleCurrencies: List<Moneda>
): CobradoActivoDetalleUi {
    val items = state.pagosActivosDetalle.map { detalle ->
        CobradoActivoDetalleItem(
            cliente = detalle.cliente,
            idPrestamo = detalle.idPrestamo,
            numeroCuota = detalle.numeroCuota,
            fechaPago = detalle.fechaPago,
            montoAbono = detalle.montoAbono,
            moneda = detalle.moneda,
            tipoCobro = detalle.tipoCobro
        )
    }.sortedByDescending { it.fechaPago }

    val totals = items
        .groupBy { it.moneda }
        .mapValues { (_, values) -> values.sumOf { it.montoAbono } }

    return CobradoActivoDetalleUi(
        title = "Cobrado activo + intereses",
        totalPrestamosActivos = state.prestamosActivosDetalle.size,
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
    stacked: Boolean = false,
    compact: Boolean = false,
    onClick: () -> Unit
) {
    val ordered = visibleCurrencies.ifEmpty { listOf(Moneda.SOLES) }
    val outerPadding = if (compact) 8.dp else 12.dp
    val innerPadding = if (compact) 6.dp else 8.dp
    val amountStyle = if (compact) MaterialTheme.typography.titleSmall else MaterialTheme.typography.titleMedium
    val nameStyle = if (compact) MaterialTheme.typography.labelSmall else MaterialTheme.typography.labelSmall
    Card(
        modifier = modifier.clickable(onClick = onClick),
        colors = androidx.compose.material3.CardDefaults.cardColors(
            containerColor = MaterialTheme.colorScheme.secondaryContainer,
            contentColor = MaterialTheme.colorScheme.onSecondaryContainer
        )
    ) {
        Column(
            modifier = Modifier.padding(outerPadding),
            verticalArrangement = Arrangement.spacedBy(if (compact) 6.dp else 8.dp)
        ) {
            DashboardCardTitle(
                title = title,
                compact = compact
            )
            val subCardContent: @Composable () -> Unit = {
                ordered.forEach { moneda ->
                    androidx.compose.material3.Surface(
                        modifier = if (stacked) Modifier.fillMaxWidth() else Modifier.weight(1f),
                        shape = RoundedCornerShape(10.dp),
                        color = MaterialTheme.colorScheme.surface.copy(alpha = 0.42f)
                    ) {
                        Column(
                            modifier = Modifier.padding(horizontal = innerPadding, vertical = innerPadding),
                            horizontalAlignment = Alignment.CenterHorizontally,
                            verticalArrangement = Arrangement.spacedBy(if (compact) 1.dp else 2.dp)
                        ) {
                            Text(
                                text = (totals[moneda] ?: 0.0).toMoney(moneda),
                                style = amountStyle,
                                color = valueColor ?: if (highlightValue) {
                                    MaterialTheme.colorScheme.tertiary
                                } else {
                                    MaterialTheme.colorScheme.onSecondaryContainer
                                },
                                maxLines = 1
                            )
                            Text(
                                text = moneda.displayName,
                                style = nameStyle,
                                color = MaterialTheme.colorScheme.onSecondaryContainer.copy(alpha = 0.72f),
                                maxLines = 1,
                                overflow = TextOverflow.Ellipsis
                            )
                        }
                    }
                }
            }
            if (stacked) {
                Column(
                    modifier = Modifier.fillMaxWidth(),
                    verticalArrangement = Arrangement.spacedBy(if (compact) 4.dp else 6.dp)
                ) {
                    subCardContent()
                }
            } else {
                Row(
                    modifier = Modifier.fillMaxWidth(),
                    horizontalArrangement = Arrangement.spacedBy(8.dp)
                ) {
                    subCardContent()
                }
            }
        }
    }
}

@Composable
private fun DashboardCardTitle(
    title: String,
    centerContent: Boolean = false,
    compact: Boolean = false
) {
    val visual = resolveDashboardTitleVisual(title)
    val iconSize = if (compact) 16.dp else 18.dp
    val textStyle = MaterialTheme.typography.labelLarge.copy(
        fontSize = if (compact) 13.sp else 14.sp,
        fontWeight = FontWeight.SemiBold
    )
    Row(
        modifier = if (centerContent) Modifier.fillMaxWidth() else Modifier,
        horizontalArrangement = if (centerContent) Arrangement.Center else Arrangement.Start,
        verticalAlignment = Alignment.CenterVertically
    ) {
        visual.icon?.let { icon ->
            Icon(
                imageVector = icon,
                contentDescription = null,
                tint = visual.color,
                modifier = Modifier.size(iconSize)
            )
            Spacer(modifier = Modifier.width(7.dp))
        }
        Text(
            text = visual.text,
            style = textStyle,
            color = visual.color,
            textAlign = if (centerContent) TextAlign.Center else TextAlign.Start
        )
    }
}

private data class DashboardTitleVisual(
    val text: String,
    val icon: ImageVector?,
    val color: Color
)

private fun resolveDashboardTitleVisual(title: String): DashboardTitleVisual = when (title) {
    "Capital prestado (activo)", "Capital prestado activo" -> DashboardTitleVisual(
        text = "Capital prestado (activo)",
        icon = Icons.Outlined.AttachMoney,
        color = Color(0xFF1B5E20)
    )

    "Total prestado activo (con intereses)", "Total activo (con intereses)", "Prestado activo + intereses" -> DashboardTitleVisual(
        text = "Total prestado activo (con intereses)",
        icon = Icons.Outlined.TrendingUp,
        color = Color(0xFF1B5E20)
    )

    "Pendiente", "Saldo pendiente" -> DashboardTitleVisual(
        text = "Pendiente",
        icon = Icons.Outlined.Schedule,
        color = Color(0xFFF57F17)
    )

    "Cobrado hoy" -> DashboardTitleVisual(
        text = "Cobrado hoy",
        icon = Icons.Outlined.Payments,
        color = Color(0xFF2E7D32)
    )

    "Cuotas vencidas" -> DashboardTitleVisual(
        text = "Cuotas vencidas",
        icon = Icons.Outlined.Warning,
        color = Color(0xFFC62828)
    )

    else -> DashboardTitleVisual(
        text = title,
        icon = null,
        color = Color(0xFF1F1F1F)
    )
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