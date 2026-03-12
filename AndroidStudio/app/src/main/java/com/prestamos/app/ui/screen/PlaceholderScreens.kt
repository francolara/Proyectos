package com.prestamos.app.ui.screen

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
import androidx.compose.material3.Button
import androidx.compose.material3.Card
import androidx.compose.material3.ExperimentalMaterial3Api
import androidx.compose.material3.ExposedDropdownMenuBox
import androidx.compose.material3.ExposedDropdownMenuDefaults
import androidx.compose.material3.ExposedDropdownMenuDefaults.TrailingIcon
import androidx.compose.material3.ExposedDropdownMenu
import androidx.compose.material3.HorizontalDivider
import androidx.compose.material3.MaterialTheme
import androidx.compose.material3.OutlinedTextField
import androidx.compose.material3.Text
import androidx.compose.runtime.Composable
import androidx.compose.runtime.getValue
import androidx.compose.runtime.mutableStateOf
import androidx.compose.runtime.remember
import androidx.compose.runtime.setValue
import androidx.compose.ui.Modifier
import androidx.compose.ui.unit.dp
import androidx.lifecycle.compose.collectAsStateWithLifecycle
import com.prestamos.app.data.local.entity.ClienteEntity
import com.prestamos.app.data.local.entity.CuotaEntity
import com.prestamos.app.data.local.entity.PrestamoEntity
import com.prestamos.app.data.local.entity.TipoPago
import com.prestamos.app.ui.viewmodel.AppViewModel
import com.prestamos.app.util.toDateString
import com.prestamos.app.util.toEpochMillis
import com.prestamos.app.util.toMoney
import java.time.LocalDate

@Composable
fun ClientesScreen(viewModel: AppViewModel) {
    val clientes by viewModel.clientes.collectAsStateWithLifecycle()
    var nombre by remember { mutableStateOf("") }
    var apellido by remember { mutableStateOf("") }
    var documento by remember { mutableStateOf("") }
    var nacionalidad by remember { mutableStateOf("") }

    LazyColumn(modifier = Modifier.fillMaxSize().padding(16.dp), verticalArrangement = Arrangement.spacedBy(10.dp)) {
        item {
            Text("Clientes", style = MaterialTheme.typography.headlineSmall)
            OutlinedTextField(nombre, { nombre = it }, label = { Text("Nombre") }, modifier = Modifier.fillMaxWidth())
            OutlinedTextField(apellido, { apellido = it }, label = { Text("Apellido") }, modifier = Modifier.fillMaxWidth())
            OutlinedTextField(documento, { documento = it }, label = { Text("Documento") }, modifier = Modifier.fillMaxWidth())
            OutlinedTextField(nacionalidad, { nacionalidad = it }, label = { Text("Nacionalidad") }, modifier = Modifier.fillMaxWidth())
            Spacer(Modifier.height(8.dp))
            Button(onClick = {
                viewModel.registrarCliente(nombre, apellido, documento, nacionalidad)
                nombre = ""; apellido = ""; documento = ""; nacionalidad = ""
            }) { Text("Guardar cliente") }
            HorizontalDivider(modifier = Modifier.padding(vertical = 12.dp))
            Text("Listado", style = MaterialTheme.typography.titleMedium)
        }
        items(clientes) { cliente ->
            Card(modifier = Modifier.fillMaxWidth()) {
                Column(Modifier.padding(12.dp)) {
                    Text("${cliente.nombre} ${cliente.apellido}")
                    Text("Doc: ${cliente.documentoIdentidad}")
                    Text("Nacionalidad: ${cliente.nacionalidad}")
                }
            }
        }
    }
}

@OptIn(ExperimentalMaterial3Api::class)
@Composable
fun PrestamosScreen(viewModel: AppViewModel) {
    val clientes by viewModel.clientes.collectAsStateWithLifecycle()
    val prestamos by viewModel.prestamos.collectAsStateWithLifecycle()

    var clienteSeleccionado by remember { mutableStateOf<ClienteEntity?>(null) }
    var monto by remember { mutableStateOf("") }
    var interes by remember { mutableStateOf("") }
    var cuotas by remember { mutableStateOf("") }
    var fechaPrimeraCuota by remember { mutableStateOf(LocalDate.now()) }
    var tipoPago by remember { mutableStateOf(TipoPago.SEMANAL) }

    var expandedCliente by remember { mutableStateOf(false) }
    var expandedTipo by remember { mutableStateOf(false) }

    LazyColumn(modifier = Modifier.fillMaxSize().padding(16.dp), verticalArrangement = Arrangement.spacedBy(10.dp)) {
        item {
            Text("Préstamos", style = MaterialTheme.typography.headlineSmall)
            ExposedDropdownMenuBox(expanded = expandedCliente, onExpandedChange = { expandedCliente = !expandedCliente }) {
                OutlinedTextField(
                    value = clienteSeleccionado?.let { "${it.nombre} ${it.apellido}" } ?: "",
                    onValueChange = {},
                    readOnly = true,
                    label = { Text("Cliente") },
                    trailingIcon = { TrailingIcon(expanded = expandedCliente) },
                    modifier = Modifier.menuAnchor().fillMaxWidth()
                )
                ExposedDropdownMenu(expanded = expandedCliente, onDismissRequest = { expandedCliente = false }) {
                    clientes.forEach { cliente ->
                        androidx.compose.material3.DropdownMenuItem(
                            text = { Text("${cliente.nombre} ${cliente.apellido}") },
                            onClick = {
                                clienteSeleccionado = cliente
                                expandedCliente = false
                            }
                        )
                    }
                }
            }
            OutlinedTextField(monto, { monto = it }, label = { Text("Monto") }, modifier = Modifier.fillMaxWidth())
            OutlinedTextField(interes, { interes = it }, label = { Text("Interés (%)") }, modifier = Modifier.fillMaxWidth())
            OutlinedTextField(cuotas, { cuotas = it }, label = { Text("Cantidad de cuotas") }, modifier = Modifier.fillMaxWidth())
            ExposedDropdownMenuBox(expanded = expandedTipo, onExpandedChange = { expandedTipo = !expandedTipo }) {
                OutlinedTextField(
                    value = tipoPago.name,
                    onValueChange = {},
                    readOnly = true,
                    label = { Text("Tipo de pago") },
                    trailingIcon = { ExposedDropdownMenuDefaults.TrailingIcon(expanded = expandedTipo) },
                    modifier = Modifier.menuAnchor().fillMaxWidth()
                )
                ExposedDropdownMenu(expanded = expandedTipo, onDismissRequest = { expandedTipo = false }) {
                    TipoPago.entries.forEach { tipo ->
                        androidx.compose.material3.DropdownMenuItem(
                            text = { Text(tipo.name) },
                            onClick = {
                                tipoPago = tipo
                                expandedTipo = false
                            }
                        )
                    }
                }
            }
            OutlinedTextField(
                value = fechaPrimeraCuota.toString(),
                onValueChange = {
                    runCatching { LocalDate.parse(it) }.onSuccess { parsed -> fechaPrimeraCuota = parsed }
                },
                label = { Text("Fecha primera cuota (yyyy-MM-dd)") },
                modifier = Modifier.fillMaxWidth()
            )
            Button(onClick = {
                clienteSeleccionado?.let {
                    viewModel.registrarPrestamo(
                        idCliente = it.idCliente,
                        monto = monto,
                        interes = interes,
                        tipoPago = tipoPago,
                        cuotas = cuotas,
                        fechaPrimeraCuota = fechaPrimeraCuota.toEpochMillis()
                    )
                    monto = ""; interes = ""; cuotas = ""
                }
            }) { Text("Guardar préstamo") }
            HorizontalDivider(modifier = Modifier.padding(vertical = 12.dp))
            Text("Listado", style = MaterialTheme.typography.titleMedium)
        }
        items(prestamos) { prestamo ->
            Card(modifier = Modifier.fillMaxWidth()) {
                Column(Modifier.padding(12.dp)) {
                    Text("Préstamo #${prestamo.idPrestamo} - Cliente ${prestamo.idCliente}")
                    Text("Total: ${prestamo.montoTotalPrestamo.toMoney()} | Cuota: ${prestamo.montoCuota.toMoney()}")
                    Text("Estado: ${prestamo.estadoPrestamo}")
                }
            }
        }
    }
}

@OptIn(ExperimentalMaterial3Api::class)
@Composable
fun PagosScreen(viewModel: AppViewModel) {
    val clientes by viewModel.clientes.collectAsStateWithLifecycle()
    val prestamos by viewModel.prestamosClientePagos.collectAsStateWithLifecycle()
    val cuotas by viewModel.cuotasPrestamoPagos.collectAsStateWithLifecycle()

    var idCliente by remember { mutableStateOf<Long?>(null) }
    var idPrestamo by remember { mutableStateOf<Long?>(null) }
    var idCuota by remember { mutableStateOf<Long?>(null) }
    var montoAbono by remember { mutableStateOf("") }
    var observacion by remember { mutableStateOf("") }

    var expandedCliente by remember { mutableStateOf(false) }
    var expandedPrestamo by remember { mutableStateOf(false) }
    var expandedCuota by remember { mutableStateOf(false) }

    Column(modifier = Modifier.fillMaxSize().padding(16.dp)) {
        Text("Pagos", style = MaterialTheme.typography.headlineSmall)

        DropdownGeneric(
            expanded = expandedCliente,
            onExpandedChange = { expandedCliente = it },
            label = "Cliente",
            selected = clientes.firstOrNull { it.idCliente == idCliente }?.let { "${it.nombre} ${it.apellido}" } ?: "",
            options = clientes,
            optionText = { "${it.nombre} ${it.apellido}" },
            onSelect = {
                idCliente = it.idCliente
                idPrestamo = null
                idCuota = null
                viewModel.seleccionarClientePagos(it.idCliente)
            }
        )

        DropdownGeneric(
            expanded = expandedPrestamo,
            onExpandedChange = { expandedPrestamo = it },
            label = "Préstamo",
            selected = prestamos.firstOrNull { it.idPrestamo == idPrestamo }?.let { "#${it.idPrestamo} - saldo ${it.montoTotalPrestamo.toMoney()}" } ?: "",
            options = prestamos,
            optionText = { "#${it.idPrestamo} - saldo ${it.montoTotalPrestamo.toMoney()}" },
            onSelect = {
                idPrestamo = it.idPrestamo
                idCuota = null
                viewModel.seleccionarPrestamoPagos(it.idPrestamo)
            }
        )

        DropdownGeneric(
            expanded = expandedCuota,
            onExpandedChange = { expandedCuota = it },
            label = "Cuota",
            selected = cuotas.firstOrNull { it.idCuota == idCuota }?.let { "Cuota ${it.numeroCuota} - pendiente ${it.saldoPendiente.toMoney()}" } ?: "",
            options = cuotas.filter { it.saldoPendiente > 0.0 },
            optionText = { "Cuota ${it.numeroCuota} - pendiente ${it.saldoPendiente.toMoney()}" },
            onSelect = { idCuota = it.idCuota }
        )

        OutlinedTextField(montoAbono, { montoAbono = it }, label = { Text("Monto abonado") }, modifier = Modifier.fillMaxWidth())
        OutlinedTextField(observacion, { observacion = it }, label = { Text("Observación") }, modifier = Modifier.fillMaxWidth())

        Button(onClick = {
            if (idPrestamo != null && idCuota != null) {
                viewModel.registrarPago(idPrestamo!!, idCuota!!, montoAbono, observacion)
                montoAbono = ""
                observacion = ""
            }
        }) { Text("Registrar pago") }
    }
}

@Composable
fun ReportesScreen(viewModel: AppViewModel) {
    val resumen by viewModel.resumenReporte.collectAsStateWithLifecycle()
    val cuotasVencidas by viewModel.cuotasVencidas.collectAsStateWithLifecycle()

    LazyColumn(modifier = Modifier.fillMaxSize().padding(16.dp), verticalArrangement = Arrangement.spacedBy(10.dp)) {
        item {
            Text("Reportes", style = MaterialTheme.typography.headlineSmall)
            Card(modifier = Modifier.fillMaxWidth()) {
                Column(Modifier.padding(12.dp)) {
                    Text("Total prestado: ${resumen.totalPrestado.toMoney()}")
                    Text("Total cobrado: ${resumen.totalCobrado.toMoney()}")
                    Text("Total pendiente: ${resumen.totalPendiente.toMoney()}")
                    Text("Préstamos activos: ${resumen.prestamosActivos}")
                    Text("Préstamos pagados: ${resumen.prestamosPagados}")
                    Text("Cuotas vencidas: ${resumen.cuotasVencidas}")
                }
            }
            Text("Detalle de cuotas vencidas", style = MaterialTheme.typography.titleMedium)
        }
        items(cuotasVencidas) { cuota ->
            Card(modifier = Modifier.fillMaxWidth()) {
                Row(Modifier.padding(12.dp), horizontalArrangement = Arrangement.SpaceBetween) {
                    Column {
                        Text("Préstamo #${cuota.idPrestamo} - Cuota ${cuota.numeroCuota}")
                        Text("Vence: ${cuota.fechaVencimiento.toDateString()}")
                    }
                    Text(cuota.saldoPendiente.toMoney())
                }
            }
        }
    }
}

@OptIn(ExperimentalMaterial3Api::class)
@Composable
private fun <T> DropdownGeneric(
    expanded: Boolean,
    onExpandedChange: (Boolean) -> Unit,
    label: String,
    selected: String,
    options: List<T>,
    optionText: (T) -> String,
    onSelect: (T) -> Unit
) {
    ExposedDropdownMenuBox(expanded = expanded, onExpandedChange = { onExpandedChange(!expanded) }) {
        OutlinedTextField(
            value = selected,
            onValueChange = {},
            readOnly = true,
            label = { Text(label) },
            trailingIcon = { ExposedDropdownMenuDefaults.TrailingIcon(expanded = expanded) },
            modifier = Modifier.menuAnchor().fillMaxWidth()
        )
        ExposedDropdownMenu(expanded = expanded, onDismissRequest = { onExpandedChange(false) }) {
            options.forEach { option ->
                androidx.compose.material3.DropdownMenuItem(
                    text = { Text(optionText(option)) },
                    onClick = {
                        onSelect(option)
                        onExpandedChange(false)
                    }
                )
            }
        }
    }
}
