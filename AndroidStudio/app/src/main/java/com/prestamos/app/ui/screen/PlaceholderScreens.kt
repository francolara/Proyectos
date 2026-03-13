package com.prestamos.app.ui.screen

import android.app.DatePickerDialog
import android.content.Intent
import android.widget.Toast
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
import androidx.compose.material3.AlertDialog
import androidx.compose.material3.Button
import androidx.compose.material3.Card
import androidx.compose.material3.DropdownMenu
import androidx.compose.material3.DropdownMenuItem
import androidx.compose.material3.ExperimentalMaterial3Api
import androidx.compose.material3.ExposedDropdownMenuBox
import androidx.compose.material3.ExposedDropdownMenuDefaults
import androidx.compose.material3.ExposedDropdownMenuDefaults.TrailingIcon
import androidx.compose.material3.HorizontalDivider
import androidx.compose.material3.MaterialTheme
import androidx.compose.material3.OutlinedTextField
import androidx.compose.material3.Text
import androidx.compose.material3.TextButton
import androidx.compose.runtime.Composable
import androidx.compose.runtime.getValue
import androidx.compose.runtime.mutableStateOf
import androidx.compose.runtime.remember
import androidx.compose.runtime.setValue
import androidx.compose.ui.Modifier
import androidx.compose.ui.platform.LocalContext
import androidx.compose.foundation.text.KeyboardOptions
import androidx.compose.ui.text.input.KeyboardType
import androidx.compose.ui.unit.dp
import androidx.core.content.FileProvider
import androidx.lifecycle.compose.collectAsStateWithLifecycle
import com.prestamos.app.data.local.entity.ClienteEntity
import com.prestamos.app.data.local.entity.EstadoPrestamo
import com.prestamos.app.data.local.entity.Moneda
import com.prestamos.app.data.local.entity.TipoPago
import com.prestamos.app.ui.screen.export.createDashboardDetalleImage
import com.prestamos.app.ui.screen.export.createDashboardDetallePdf
import com.prestamos.app.ui.viewmodel.AppViewModel
import com.prestamos.app.util.toDateString
import com.prestamos.app.util.toEpochMillis
import com.prestamos.app.util.toMoney
import java.io.File
import java.time.LocalDate

@Composable
fun ClientesScreen(viewModel: AppViewModel) {
    val clientes by viewModel.clientes.collectAsStateWithLifecycle()
    var nombre by remember { mutableStateOf("") }
    var apellido by remember { mutableStateOf("") }
    var documento by remember { mutableStateOf("") }
    var direccion by remember { mutableStateOf("") }
    var telefono by remember { mutableStateOf("") }

    LazyColumn(
        modifier = Modifier
            .fillMaxSize()
            .padding(16.dp),
        verticalArrangement = Arrangement.spacedBy(10.dp)
    ) {
        item {
            Text("Clientes", style = MaterialTheme.typography.headlineSmall)
            OutlinedTextField(
                value = nombre,
                onValueChange = { nombre = sanitizeSingleLine(it) },
                label = { Text("Nombres *") },
                singleLine = true,
                modifier = Modifier.fillMaxWidth()
            )
            OutlinedTextField(
                value = apellido,
                onValueChange = { apellido = sanitizeSingleLine(it) },
                label = { Text("Apellido *") },
                singleLine = true,
                modifier = Modifier.fillMaxWidth()
            )
            OutlinedTextField(
                value = documento,
                onValueChange = { documento = onlyDigits(it) },
                label = { Text("Documento de identidad *") },
                singleLine = true,
                keyboardOptions = KeyboardOptions(keyboardType = KeyboardType.Number),
                modifier = Modifier.fillMaxWidth()
            )
            OutlinedTextField(
                value = direccion,
                onValueChange = { direccion = sanitizeSingleLine(it) },
                label = { Text("Dirección") },
                singleLine = true,
                modifier = Modifier.fillMaxWidth()
            )
            OutlinedTextField(
                value = telefono,
                onValueChange = { telefono = onlyDigits(it) },
                label = { Text("Nro de teléfono") },
                singleLine = true,
                keyboardOptions = KeyboardOptions(keyboardType = KeyboardType.Phone),
                modifier = Modifier.fillMaxWidth()
            )
            Spacer(Modifier.height(8.dp))
            Button(onClick = {
                viewModel.registrarCliente(nombre, apellido, documento, direccion, telefono)
                nombre = ""
                apellido = ""
                documento = ""
                direccion = ""
                telefono = ""
            }) { Text("Guardar cliente") }
            HorizontalDivider(modifier = Modifier.padding(vertical = 12.dp))
            Text("Listado", style = MaterialTheme.typography.titleMedium)
        }

        items(clientes) { cliente ->
            Card(modifier = Modifier.fillMaxWidth()) {
                Column(Modifier.padding(12.dp)) {
                    Text("${cliente.nombre} ${cliente.apellido}")
                    Text("Doc: ${cliente.documentoIdentidad}")
                    Text("Dirección: ${cliente.direccion.ifBlank { "-" }}")
                    Text("Teléfono: ${cliente.telefono.ifBlank { "-" }}")
                }
            }
        }
    }
}

private enum class PrestamosFiltroEstado(val estado: EstadoPrestamo, val label: String) {
    ACTIVOS(EstadoPrestamo.ACTIVO, "ACTIVO"),
    PAGADOS(EstadoPrestamo.PAGADO, "PAGADO")
}

@OptIn(ExperimentalMaterial3Api::class)
@Composable
fun PrestamosScreen(viewModel: AppViewModel) {
    val context = LocalContext.current
    val clientes by viewModel.clientes.collectAsStateWithLifecycle()
    val prestamos by viewModel.prestamos.collectAsStateWithLifecycle()

    var filtroCliente by remember { mutableStateOf("") }
    val clientesFiltrados = remember(clientes, filtroCliente) {
        filtrarClientes(clientes, filtroCliente)
    }

    var clienteSeleccionado by remember { mutableStateOf<ClienteEntity?>(null) }
    var monto by remember { mutableStateOf("") }
    var interes by remember { mutableStateOf("") }
    var cuotas by remember { mutableStateOf("") }
    var fechaPrimeraCuota by remember { mutableStateOf(LocalDate.now()) }
    var moneda by remember { mutableStateOf(Moneda.SOLES) }
    var tipoPago by remember { mutableStateOf(TipoPago.SEMANAL) }

    var expandedCliente by remember { mutableStateOf(false) }
    var expandedMoneda by remember { mutableStateOf(false) }
    var expandedTipo by remember { mutableStateOf(false) }
    var mostrarDetallePrestamo by remember { mutableStateOf(false) }
    var prestamoDetalleId by remember { mutableStateOf<Long?>(null) }
    var filtroEstado by remember { mutableStateOf(PrestamosFiltroEstado.ACTIVOS) }
    var expandedFiltroEstado by remember { mutableStateOf(false) }

    val prestamosFiltrados = remember(prestamos, filtroEstado) {
        prestamos.filter { it.estadoPrestamo == filtroEstado.estado }
    }

    val cuotasDetalle by viewModel.cuotasPrestamoDetalle.collectAsStateWithLifecycle()

    LazyColumn(
        modifier = Modifier
            .fillMaxSize()
            .padding(16.dp),
        verticalArrangement = Arrangement.spacedBy(10.dp)
    ) {
        item {
            Text("Préstamos", style = MaterialTheme.typography.headlineSmall)

            OutlinedTextField(
                value = filtroCliente,
                onValueChange = { filtroCliente = sanitizeSingleLine(it) },
                label = { Text("Filtrar cliente (nombre o documento)") },
                singleLine = true,
                modifier = Modifier.fillMaxWidth()
            )

            ExposedDropdownMenuBox(expanded = expandedCliente, onExpandedChange = { expandedCliente = !expandedCliente }) {
                OutlinedTextField(
                    value = clienteSeleccionado?.let { "${it.nombre} ${it.apellido}" } ?: "",
                    onValueChange = {},
                    readOnly = true,
                    label = { Text("Cliente") },
                    trailingIcon = { TrailingIcon(expanded = expandedCliente) },
                    modifier = Modifier.menuAnchor().fillMaxWidth()
                )
                DropdownMenu(expanded = expandedCliente, onDismissRequest = { expandedCliente = false }) {
                    clientesFiltrados.forEach { cliente ->
                        DropdownMenuItem(
                            text = { Text("${cliente.nombre} ${cliente.apellido} (${cliente.documentoIdentidad})") },
                            onClick = {
                                clienteSeleccionado = cliente
                                expandedCliente = false
                            }
                        )
                    }
                }
            }

            OutlinedTextField(
                value = monto,
                onValueChange = { monto = onlyDecimal(it) },
                label = { Text("Monto") },
                singleLine = true,
                keyboardOptions = KeyboardOptions(keyboardType = KeyboardType.Decimal),
                modifier = Modifier.fillMaxWidth()
            )
            OutlinedTextField(
                value = interes,
                onValueChange = { interes = onlyDecimal(it) },
                label = { Text("Interés (%)") },
                singleLine = true,
                keyboardOptions = KeyboardOptions(keyboardType = KeyboardType.Decimal),
                modifier = Modifier.fillMaxWidth()
            )

            ExposedDropdownMenuBox(expanded = expandedMoneda, onExpandedChange = { expandedMoneda = !expandedMoneda }) {
                OutlinedTextField(
                    value = when (moneda) {
                        Moneda.SOLES -> "Soles"
                        Moneda.DOLARES -> "Dólares"
                    },
                    onValueChange = {},
                    readOnly = true,
                    label = { Text("Moneda") },
                    trailingIcon = { ExposedDropdownMenuDefaults.TrailingIcon(expanded = expandedMoneda) },
                    modifier = Modifier.menuAnchor().fillMaxWidth()
                )
                DropdownMenu(expanded = expandedMoneda, onDismissRequest = { expandedMoneda = false }) {
                    Moneda.entries.forEach { opcion ->
                        DropdownMenuItem(
                            text = { Text(if (opcion == Moneda.SOLES) "Soles" else "Dólares") },
                            onClick = {
                                moneda = opcion
                                expandedMoneda = false
                            }
                        )
                    }
                }
            }

            OutlinedTextField(
                value = cuotas,
                onValueChange = { cuotas = onlyDigits(it) },
                label = { Text("Cantidad de cuotas") },
                singleLine = true,
                keyboardOptions = KeyboardOptions(keyboardType = KeyboardType.Number),
                modifier = Modifier.fillMaxWidth()
            )

            ExposedDropdownMenuBox(expanded = expandedTipo, onExpandedChange = { expandedTipo = !expandedTipo }) {
                OutlinedTextField(
                    value = tipoPago.name,
                    onValueChange = {},
                    readOnly = true,
                    label = { Text("Tipo de pago") },
                    trailingIcon = { ExposedDropdownMenuDefaults.TrailingIcon(expanded = expandedTipo) },
                    modifier = Modifier.menuAnchor().fillMaxWidth()
                )
                DropdownMenu(expanded = expandedTipo, onDismissRequest = { expandedTipo = false }) {
                    TipoPago.entries.forEach { tipo ->
                        DropdownMenuItem(
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
                onValueChange = {},
                readOnly = true,
                label = { Text("Fecha primera cuota") },
                singleLine = true,
                modifier = Modifier.fillMaxWidth()
            )
            Button(onClick = {
                DatePickerDialog(
                    context,
                    { _, year, month, dayOfMonth ->
                        fechaPrimeraCuota = LocalDate.of(year, month + 1, dayOfMonth)
                    },
                    fechaPrimeraCuota.year,
                    fechaPrimeraCuota.monthValue - 1,
                    fechaPrimeraCuota.dayOfMonth
                ).show()
            }) {
                Text("Seleccionar fecha")
            }

            Button(onClick = {
                clienteSeleccionado?.let {
                    viewModel.registrarPrestamo(
                        idCliente = it.idCliente,
                        monto = monto,
                        interes = interes,
                        moneda = moneda,
                        tipoPago = tipoPago,
                        cuotas = cuotas,
                        fechaPrimeraCuota = fechaPrimeraCuota.toEpochMillis()
                    )
                    monto = ""
                    interes = ""
                    cuotas = ""
                }
            }) { Text("Guardar préstamo") }

            HorizontalDivider(modifier = Modifier.padding(vertical = 12.dp))
            Text("Listado", style = MaterialTheme.typography.titleMedium)

            ExposedDropdownMenuBox(
                expanded = expandedFiltroEstado,
                onExpandedChange = { expandedFiltroEstado = !expandedFiltroEstado }
            ) {
                OutlinedTextField(
                    value = filtroEstado.label,
                    onValueChange = {},
                    readOnly = true,
                    label = { Text("Filtro estado") },
                    trailingIcon = { ExposedDropdownMenuDefaults.TrailingIcon(expanded = expandedFiltroEstado) },
                    modifier = Modifier.menuAnchor().fillMaxWidth()
                )
                DropdownMenu(expanded = expandedFiltroEstado, onDismissRequest = { expandedFiltroEstado = false }) {
                    PrestamosFiltroEstado.entries.forEach { opcion ->
                        DropdownMenuItem(
                            text = { Text(opcion.label) },
                            onClick = {
                                filtroEstado = opcion
                                expandedFiltroEstado = false
                            }
                        )
                    }
                }
            }
        }

        items(prestamosFiltrados) { prestamo ->
            val cliente = clientes.firstOrNull { it.idCliente == prestamo.idCliente }
            Card(
                modifier = Modifier
                    .fillMaxWidth()
                    .clickable {
                        prestamoDetalleId = prestamo.idPrestamo
                        viewModel.seleccionarPrestamoDetalle(prestamo.idPrestamo)
                        mostrarDetallePrestamo = true
                    }
            ) {
                Column(Modifier.padding(12.dp)) {
                    Text("Préstamo #${prestamo.idPrestamo}")
                    Text("Cliente: ${cliente?.nombre ?: "-"} ${cliente?.apellido ?: ""}".trim())
                    Text(
                        "Total: ${prestamo.montoTotalPrestamo.toMoney(prestamo.moneda)} | " +
                            "Cuota: ${prestamo.montoCuota.toMoney(prestamo.moneda)}"
                    )
                    Text("Moneda: ${if (prestamo.moneda == Moneda.SOLES) "Soles" else "Dólares"}")
                    Text("Estado: ${prestamo.estadoPrestamo}")
                }
            }
        }
    }

    if (mostrarDetallePrestamo && prestamoDetalleId != null) {
        val prestamo = prestamos.firstOrNull { it.idPrestamo == prestamoDetalleId }
        val cliente = clientes.firstOrNull { it.idCliente == prestamo?.idCliente }
        val moneda = prestamo?.moneda ?: Moneda.SOLES
        val totalDeudaPendiente = cuotasDetalle.sumOf { it.saldoPendiente }
        val cronograma = if (cuotasDetalle.isEmpty()) {
            "Sin cuotas registradas"
        } else {
            cuotasDetalle.joinToString("\n") { cuota ->
                "Cuota ${cuota.numeroCuota}: vence ${cuota.fechaVencimiento.toDateString()} | " +
                    "monto ${cuota.montoCuota.toMoney(moneda)} | " +
                    "pendiente ${cuota.saldoPendiente.toMoney(moneda)}"
            }
        }
        val detallePrestamo = buildString {
            appendLine("Detalle del préstamo")
            appendLine("Cliente: ${cliente?.nombre ?: "-"} ${cliente?.apellido ?: ""}".trim())
            appendLine("Préstamo #${prestamo?.idPrestamo ?: "-"}")
            appendLine("Monto prestado: ${prestamo?.montoPrestado?.toMoney(moneda) ?: "-"}")
            appendLine("Monto total: ${prestamo?.montoTotalPrestamo?.toMoney(moneda) ?: "-"}")
            appendLine("Interés: ${prestamo?.interes ?: "-"}%")
            appendLine("Tipo pago: ${prestamo?.tipoPago ?: "-"}")
            appendLine("Cuotas: ${prestamo?.cantidadCuotas ?: "-"}")
            appendLine("Fecha registro préstamo: ${prestamo?.fechaRegistro?.toDateString() ?: "-"}")
            appendLine("\nCronograma de cuotas")
            appendLine(cronograma)
            appendLine("\nTotal deuda pendiente: ${totalDeudaPendiente.toMoney(moneda)}")
        }

        AlertDialog(
            onDismissRequest = {
                mostrarDetallePrestamo = false
                viewModel.seleccionarPrestamoDetalle(null)
            },
            confirmButton = {
                TextButton(onClick = {
                    compartirTexto(context, "Detalle préstamo", detallePrestamo)
                }) {
                    Text("Compartir")
                }
            },
            dismissButton = {
                Row(horizontalArrangement = Arrangement.spacedBy(8.dp)) {
                    TextButton(onClick = {
                        runCatching {
                            createDashboardDetalleImage(context, "Detalle préstamo", detallePrestamo)
                        }.onSuccess { file ->
                            compartirArchivo(context, file, "image/png")
                        }.onFailure {
                            Toast.makeText(context, "No se pudo exportar imagen", Toast.LENGTH_SHORT).show()
                        }
                    }) {
                        Text("Imagen")
                    }
                    TextButton(onClick = {
                        runCatching {
                            createDashboardDetallePdf(context, "Detalle préstamo", detallePrestamo)
                        }.onSuccess { file ->
                            compartirArchivo(context, file, "application/pdf")
                        }.onFailure {
                            Toast.makeText(context, "No se pudo exportar PDF", Toast.LENGTH_SHORT).show()
                        }
                    }) {
                        Text("PDF")
                    }
                    TextButton(onClick = {
                        mostrarDetallePrestamo = false
                        viewModel.seleccionarPrestamoDetalle(null)
                    }) {
                        Text("Cerrar")
                    }
                }
            },
            title = { Text("Detalle del préstamo") },
            text = {
                LazyColumn(verticalArrangement = Arrangement.spacedBy(6.dp)) {
                    item {
                        Text("Cliente: ${cliente?.nombre ?: "-"} ${cliente?.apellido ?: ""}".trim())
                        Text("Préstamo #${prestamo?.idPrestamo ?: "-"}")
                        Text("Monto prestado: ${prestamo?.montoPrestado?.toMoney(moneda) ?: "-"}")
                        Text("Monto total: ${prestamo?.montoTotalPrestamo?.toMoney(moneda) ?: "-"}")
                        Text("Interés: ${prestamo?.interes ?: "-"}%")
                        Text("Tipo pago: ${prestamo?.tipoPago ?: "-"}")
                        Text("Cuotas: ${prestamo?.cantidadCuotas ?: "-"}")
                        Text("Fecha registro préstamo: ${prestamo?.fechaRegistro?.toDateString() ?: "-"}")
                    }
                    item { HorizontalDivider() }
                    item { Text("Cronograma de cuotas", style = MaterialTheme.typography.titleSmall) }
                    items(cuotasDetalle) { cuota ->
                        Text(
                            "Cuota ${cuota.numeroCuota}: vence ${cuota.fechaVencimiento.toDateString()} | " +
                                "monto ${cuota.montoCuota.toMoney(moneda)} | " +
                                "pendiente ${cuota.saldoPendiente.toMoney(moneda)}"
                        )
                    }
                    item {
                        HorizontalDivider()
                        Text(
                            "Total deuda pendiente: ${totalDeudaPendiente.toMoney(moneda)}",
                            style = MaterialTheme.typography.titleSmall
                        )
                    }
                }
            }
        )
    }

}

@OptIn(ExperimentalMaterial3Api::class)
@Composable
fun PagosScreen(viewModel: AppViewModel) {
    val clientes by viewModel.clientes.collectAsStateWithLifecycle()
    val prestamos by viewModel.prestamosClientePagos.collectAsStateWithLifecycle()
    val cuotas by viewModel.cuotasPrestamoPagos.collectAsStateWithLifecycle()

    var filtroCliente by remember { mutableStateOf("") }
    val clientesFiltrados = remember(clientes, filtroCliente) {
        filtrarClientes(clientes, filtroCliente)
    }

    var idCliente by remember { mutableStateOf<Long?>(null) }
    var idPrestamo by remember { mutableStateOf<Long?>(null) }
    var idCuota by remember { mutableStateOf<Long?>(null) }
    var montoAbono by remember { mutableStateOf("") }
    var observacion by remember { mutableStateOf("") }

    var expandedCliente by remember { mutableStateOf(false) }
    var expandedPrestamo by remember { mutableStateOf(false) }
    var expandedCuota by remember { mutableStateOf(false) }

    Column(
        modifier = Modifier
            .fillMaxSize()
            .padding(16.dp)
    ) {
        Text("Pagos", style = MaterialTheme.typography.headlineSmall)

        OutlinedTextField(
            value = filtroCliente,
            onValueChange = { filtroCliente = sanitizeSingleLine(it) },
            label = { Text("Filtrar cliente (nombre o documento)") },
            singleLine = true,
            modifier = Modifier.fillMaxWidth()
        )

        DropdownGeneric(
            expanded = expandedCliente,
            onExpandedChange = { expandedCliente = it },
            label = "Cliente",
            selected = clientes.firstOrNull { it.idCliente == idCliente }?.let { "${it.nombre} ${it.apellido}" } ?: "",
            options = clientesFiltrados,
            optionText = { "${it.nombre} ${it.apellido} (${it.documentoIdentidad})" },
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
            selected = prestamos.firstOrNull { it.idPrestamo == idPrestamo }
                ?.let { "#${it.idPrestamo} - saldo ${it.montoTotalPrestamo.toMoney(it.moneda)}" } ?: "",
            options = prestamos,
            optionText = { "#${it.idPrestamo} - saldo ${it.montoTotalPrestamo.toMoney(it.moneda)}" },
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
            selected = cuotas.firstOrNull { it.idCuota == idCuota }
                ?.let { "Cuota ${it.numeroCuota} - pendiente ${it.saldoPendiente.toMoney()}" } ?: "",
            options = cuotas.filter { it.saldoPendiente > 0.0 },
            optionText = { "Cuota ${it.numeroCuota} - pendiente ${it.saldoPendiente.toMoney()}" },
            onSelect = { idCuota = it.idCuota }
        )

        OutlinedTextField(
            value = montoAbono,
            onValueChange = { montoAbono = onlyDecimal(it) },
            label = { Text("Monto abonado") },
            singleLine = true,
            keyboardOptions = KeyboardOptions(keyboardType = KeyboardType.Decimal),
            modifier = Modifier.fillMaxWidth()
        )
        OutlinedTextField(
            value = observacion,
            onValueChange = { observacion = sanitizeSingleLine(it) },
            label = { Text("Observación") },
            singleLine = true,
            modifier = Modifier.fillMaxWidth()
        )

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
    val clientes by viewModel.clientes.collectAsStateWithLifecycle()
    val prestamos by viewModel.prestamos.collectAsStateWithLifecycle()

    LazyColumn(
        modifier = Modifier
            .fillMaxSize()
            .padding(16.dp),
        verticalArrangement = Arrangement.spacedBy(10.dp)
    ) {
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
            val prestamo = prestamos.firstOrNull { it.idPrestamo == cuota.idPrestamo }
            val cliente = clientes.firstOrNull { it.idCliente == prestamo?.idCliente }
            Card(modifier = Modifier.fillMaxWidth()) {
                Row(Modifier.padding(12.dp), horizontalArrangement = Arrangement.SpaceBetween) {
                    Column {
                        Text("Cliente: ${cliente?.nombre ?: "-"} ${cliente?.apellido ?: ""}".trim())
                        Text("Préstamo #${cuota.idPrestamo} - Cuota ${cuota.numeroCuota}")
                        Text("Registro préstamo: ${prestamo?.fechaRegistro?.toDateString() ?: "-"}")
                        Text("Vence: ${cuota.fechaVencimiento.toDateString()}")
                    }
                    Text(cuota.saldoPendiente.toMoney(prestamo?.moneda ?: Moneda.SOLES))
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
        DropdownMenu(expanded = expanded, onDismissRequest = { onExpandedChange(false) }) {
            options.forEach { option ->
                DropdownMenuItem(
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


private fun compartirTexto(context: android.content.Context, titulo: String, detalle: String) {
    val sendIntent = Intent(Intent.ACTION_SEND).apply {
        type = "text/plain"
        putExtra(Intent.EXTRA_SUBJECT, titulo)
        putExtra(Intent.EXTRA_TEXT, detalle)
    }
    context.startActivity(Intent.createChooser(sendIntent, "Compartir detalle"))
}

private fun compartirArchivo(context: android.content.Context, file: File, mimeType: String) {
    val uri = FileProvider.getUriForFile(context, "${context.packageName}.fileprovider", file)
    val sendIntent = Intent(Intent.ACTION_SEND).apply {
        type = mimeType
        putExtra(Intent.EXTRA_STREAM, uri)
        putExtra(Intent.EXTRA_SUBJECT, "Detalle préstamo")
        putExtra(Intent.EXTRA_TEXT, "Detalle exportado")
        addFlags(Intent.FLAG_GRANT_READ_URI_PERMISSION)
    }
    context.startActivity(Intent.createChooser(sendIntent, "Compartir detalle"))
}

private fun filtrarClientes(clientes: List<ClienteEntity>, filtro: String): List<ClienteEntity> {
    val query = filtro.trim().lowercase()
    if (query.isBlank()) return clientes
    return clientes.filter { cliente ->
        "${cliente.nombre} ${cliente.apellido}".lowercase().contains(query) ||
            cliente.documentoIdentidad.lowercase().contains(query)
    }
}

private fun sanitizeSingleLine(value: String): String = value.filter { it != '\n' && it != '\r' }

private fun onlyDigits(value: String): String = value.filter { it.isDigit() }

private fun onlyDecimal(value: String): String {
    val clean = sanitizeSingleLine(value).replace(',', '.')
    var hasDot = false
    val result = StringBuilder()
    clean.forEach { ch ->
        when {
            ch.isDigit() -> result.append(ch)
            ch == '.' && !hasDot -> {
                hasDot = true
                result.append(ch)
            }
        }
    }
    return result.toString()
}
