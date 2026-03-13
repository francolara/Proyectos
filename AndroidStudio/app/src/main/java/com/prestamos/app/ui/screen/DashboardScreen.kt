package com.prestamos.app.ui.screen

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
import androidx.compose.material.icons.Icons
import androidx.compose.material.icons.outlined.Logout
import androidx.compose.material3.Card
import androidx.compose.material3.Icon
import androidx.compose.material3.IconButton
import androidx.compose.material3.MaterialTheme
import androidx.compose.material3.Text
import androidx.compose.runtime.Composable
import androidx.compose.runtime.getValue
import androidx.compose.runtime.mutableStateOf
import androidx.compose.runtime.remember
import androidx.compose.runtime.setValue
import androidx.compose.ui.Alignment
import androidx.compose.ui.Modifier
import androidx.compose.ui.graphics.Color
import androidx.compose.ui.unit.dp
import androidx.lifecycle.compose.collectAsStateWithLifecycle
import com.prestamos.app.navigation.AppDestinations
import com.prestamos.app.ui.viewmodel.DashboardViewModel
import com.prestamos.app.util.toDateString
import com.prestamos.app.util.toMoney

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
    onLogout: () -> Unit,
    onNavigate: (String) -> Unit
) {
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
                IconButton(onClick = onLogout) {
                    Icon(Icons.Outlined.Logout, contentDescription = "Cerrar sesión")
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
                DashboardCard("Pagadas", state.estadoCuotas["Pagadas"].toString(), Modifier.weight(1f)) { detalleSeleccionado = DashboardDetalle.CUOTAS_PAGADAS }
                DashboardCard("Pendientes", state.estadoCuotas["Pendientes"].toString(), Modifier.weight(1f)) { detalleSeleccionado = DashboardDetalle.CUOTAS_PENDIENTES }
            }
            Spacer(Modifier.height(8.dp))
            Row(modifier = Modifier.fillMaxWidth(), horizontalArrangement = Arrangement.spacedBy(8.dp)) {
                DashboardCard("Parciales", state.estadoCuotas["Parciales"].toString(), Modifier.weight(1f)) { detalleSeleccionado = DashboardDetalle.CUOTAS_PARCIALES }
                DashboardCard("Vencidas", state.estadoCuotas["Vencidas"].toString(), Modifier.weight(1f)) { detalleSeleccionado = DashboardDetalle.CUOTAS_VENCIDAS }
            }
        }

        item {
            Text("Próximos vencimientos", style = MaterialTheme.typography.titleMedium)
            if (state.proximosVencimientos.isEmpty()) {
                Text("No hay próximos vencimientos")
            }
        }
        items(state.proximosVencimientos) { cuota ->
            Card(modifier = Modifier.fillMaxWidth().clickable { onNavigate(AppDestinations.PRESTAMOS.route) }) {
                Column(Modifier.padding(10.dp)) {
                    Text(cuota.cliente)
                    Text("Cuota ${cuota.numeroCuota} - Vence ${cuota.fechaVencimiento.toDateString()}")
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
            Card(modifier = Modifier.fillMaxWidth().clickable { onNavigate(AppDestinations.PAGOS.route) }) {
                Column(Modifier.padding(10.dp)) {
                    Text(pago.cliente)
                    Text("Fecha: ${pago.fechaPago.toDateString()}")
                    Text("Abono: ${pago.montoAbono.toMoney(pago.moneda)}")
                }
            }
        }

        item {
            Text("Accesos rápidos", style = MaterialTheme.typography.titleMedium)
            Row(modifier = Modifier.fillMaxWidth(), horizontalArrangement = Arrangement.spacedBy(8.dp)) {
                DashboardCard("Nuevo cliente", "Ir", Modifier.weight(1f)) { onNavigate(AppDestinations.CLIENTES.route) }
                DashboardCard("Nuevo préstamo", "Ir", Modifier.weight(1f)) { onNavigate(AppDestinations.PRESTAMOS.route) }
            }
            Spacer(Modifier.height(8.dp))
            Row(modifier = Modifier.fillMaxWidth(), horizontalArrangement = Arrangement.spacedBy(8.dp)) {
                DashboardCard("Registrar pago", "Ir", Modifier.weight(1f)) { onNavigate(AppDestinations.PAGOS.route) }
                DashboardCard("Ver reportes", "Ir", Modifier.weight(1f)) { onNavigate(AppDestinations.REPORTES.route) }
            }
        }
    }

    detalleSeleccionado?.let {
        DetalleDashboardDialog(it, onClose = { detalleSeleccionado = null })
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
private fun DetalleDashboardDialog(detalle: DashboardDetalle, onClose: () -> Unit) {
    androidx.compose.material3.AlertDialog(
        onDismissRequest = onClose,
        confirmButton = {
            androidx.compose.material3.TextButton(onClick = onClose) {
                Text("Cerrar")
            }
        },
        title = { Text("Detalle") },
        text = {
            Text(
                when (detalle) {
                    DashboardDetalle.CAPITAL -> "Detalle de préstamos activos"
                    DashboardDetalle.PENDIENTE -> "Detalle de cuotas pendientes"
                    DashboardDetalle.COBRADO_HOY -> "Detalle de pagos de hoy"
                    DashboardDetalle.VENCIDAS -> "Detalle de cuotas vencidas"
                    DashboardDetalle.CUOTAS_PAGADAS -> "Detalle de cuotas pagadas"
                    DashboardDetalle.CUOTAS_PENDIENTES -> "Detalle de cuotas pendientes"
                    DashboardDetalle.CUOTAS_PARCIALES -> "Detalle de cuotas parciales"
                    DashboardDetalle.CUOTAS_VENCIDAS -> "Detalle de cuotas vencidas"
                }
            )
        }
    )
}
