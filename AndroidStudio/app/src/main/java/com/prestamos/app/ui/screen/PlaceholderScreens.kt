package com.prestamos.app.ui.screen

import android.app.DatePickerDialog
import android.content.Intent
import android.widget.Toast
import androidx.compose.foundation.clickable
import androidx.compose.foundation.gestures.detectTapGestures
import androidx.compose.foundation.layout.Arrangement
import androidx.compose.foundation.layout.Box
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
import androidx.compose.material.icons.Icons
import androidx.compose.material.icons.outlined.Close
import androidx.compose.material.icons.outlined.DateRange
import androidx.compose.material.icons.outlined.Search
import androidx.compose.material3.Icon
import androidx.compose.material3.IconButton
import androidx.compose.material3.MaterialTheme
import androidx.compose.material3.OutlinedTextField
import androidx.compose.material3.Text
import androidx.compose.material3.TextButton
import androidx.compose.runtime.Composable
import androidx.compose.runtime.LaunchedEffect
import androidx.compose.runtime.getValue
import androidx.compose.runtime.mutableStateOf
import androidx.compose.runtime.remember
import androidx.compose.runtime.setValue
import androidx.compose.ui.Modifier
import androidx.compose.ui.input.pointer.pointerInput
import androidx.compose.ui.platform.LocalContext
import androidx.compose.foundation.text.KeyboardOptions
import androidx.compose.ui.text.input.KeyboardType
import androidx.compose.ui.unit.dp
import androidx.compose.ui.window.PopupProperties
import androidx.core.content.FileProvider
import androidx.lifecycle.compose.collectAsStateWithLifecycle
import com.prestamos.app.data.local.entity.ClienteEntity
import com.prestamos.app.data.local.entity.EstadoPrestamo
import com.prestamos.app.data.local.entity.Moneda
import com.prestamos.app.data.local.entity.TipoPago
import com.prestamos.app.ui.screen.export.createDashboardDetallePdf
import com.prestamos.app.ui.viewmodel.AppViewModel
import com.prestamos.app.util.toDateString
import com.prestamos.app.util.toEpochMillis
import com.prestamos.app.util.toMoney
import java.io.File
import java.time.Instant
import java.time.LocalDate
import java.time.ZoneId
import java.time.temporal.ChronoUnit

// Firma Codex 2026-03-17

@Composable
fun ClientesScreen(viewModel: AppViewModel) {
    val clientes by viewModel.clientes.collectAsStateWithLifecycle()
    var nombre by remember { mutableStateOf("") }
    var apellido by remember { mutableStateOf("") }
    var documento by remember { mutableStateOf("") }
    var direccion by remember { mutableStateOf("") }
    var telefono by remember { mutableStateOf("") }
    var mostrarRegistroOk by remember { mutableStateOf(false) }
    var clienteEditando by remember { mutableStateOf<ClienteEntity?>(null) }
    var nombreEdit by remember { mutableStateOf("") }
    var apellidoEdit by remember { mutableStateOf("") }
    var direccionEdit by remember { mutableStateOf("") }
    var telefonoEdit by remember { mutableStateOf("") }
    var clienteAEliminar by remember { mutableStateOf<ClienteEntity?>(null) }

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
                label = { Text("Direccion") },
                singleLine = true,
                modifier = Modifier.fillMaxWidth()
            )
            OutlinedTextField(
                value = telefono,
                onValueChange = { telefono = onlyDigits(it) },
                label = { Text("Nro de telefono") },
                singleLine = true,
                keyboardOptions = KeyboardOptions(keyboardType = KeyboardType.Phone),
                modifier = Modifier.fillMaxWidth()
            )
            Spacer(Modifier.height(8.dp))
            Button(onClick = {
                viewModel.registrarCliente(
                    nombre = nombre,
                    apellido = apellido,
                    documento = documento,
                    direccion = direccion,
                    telefono = telefono
                ) {
                    nombre = ""
                    apellido = ""
                    documento = ""
                    direccion = ""
                    telefono = ""
                    mostrarRegistroOk = true
                }
            }) { Text("Guardar cliente") }
            HorizontalDivider(modifier = Modifier.padding(vertical = 12.dp))
            Text("Listado", style = MaterialTheme.typography.titleMedium)
        }

        items(clientes) { cliente ->
            Card(modifier = Modifier.fillMaxWidth()) {
                Column(Modifier.padding(12.dp), verticalArrangement = Arrangement.spacedBy(4.dp)) {
                    Text("\uD83D\uDC64 ${cliente.nombre} ${cliente.apellido} / \uD83D\uDCC4 ${cliente.documentoIdentidad}")
                    Text("\uD83D\uDCCD ${cliente.direccion.ifBlank { "-" }}")
                    Text("\uD83D\uDCDE ${cliente.telefono.ifBlank { "-" }}")
                    Row(horizontalArrangement = Arrangement.spacedBy(12.dp)) {
                        TextButton(onClick = {
                            clienteEditando = cliente
                            nombreEdit = cliente.nombre
                            apellidoEdit = cliente.apellido
                            direccionEdit = cliente.direccion
                            telefonoEdit = cliente.telefono
                        }) {
                            Text("\u270F\uFE0F Editar")
                        }
                        TextButton(onClick = { clienteAEliminar = cliente }) {
                            Text("\uD83D\uDDD1 Eliminar")
                        }
                    }
                }
            }
        }
    }

    if (mostrarRegistroOk) {
        AlertDialog(
            onDismissRequest = { mostrarRegistroOk = false },
            confirmButton = {
                TextButton(onClick = { mostrarRegistroOk = false }) { Text("Aceptar") }
            },
            title = { Text("Registro") },
            text = { Text("Se realizo el registro correctamente") }
        )
    }

    if (clienteEditando != null) {
        AlertDialog(
            onDismissRequest = { clienteEditando = null },
            confirmButton = {
                TextButton(onClick = {
                    val cliente = clienteEditando ?: return@TextButton
                    viewModel.actualizarCliente(
                        idCliente = cliente.idCliente,
                        nombre = nombreEdit,
                        apellido = apellidoEdit,
                        direccion = direccionEdit,
                        telefono = telefonoEdit
                    )
                    clienteEditando = null
                }) { Text("Guardar") }
            },
            dismissButton = {
                TextButton(onClick = { clienteEditando = null }) { Text("Cancelar") }
            },
            title = { Text("Actualizar cliente") },
            text = {
                Column(verticalArrangement = Arrangement.spacedBy(8.dp)) {
                    OutlinedTextField(
                        value = nombreEdit,
                        onValueChange = { nombreEdit = sanitizeSingleLine(it) },
                        label = { Text("Nombres *") },
                        singleLine = true,
                        modifier = Modifier.fillMaxWidth()
                    )
                    OutlinedTextField(
                        value = apellidoEdit,
                        onValueChange = { apellidoEdit = sanitizeSingleLine(it) },
                        label = { Text("Apellido *") },
                        singleLine = true,
                        modifier = Modifier.fillMaxWidth()
                    )
                    OutlinedTextField(
                        value = direccionEdit,
                        onValueChange = { direccionEdit = sanitizeSingleLine(it) },
                        label = { Text("Direccion") },
                        singleLine = true,
                        modifier = Modifier.fillMaxWidth()
                    )
                    OutlinedTextField(
                        value = telefonoEdit,
                        onValueChange = { telefonoEdit = onlyDigits(it) },
                        label = { Text("Telefono") },
                        singleLine = true,
                        keyboardOptions = KeyboardOptions(keyboardType = KeyboardType.Phone),
                        modifier = Modifier.fillMaxWidth()
                    )
                }
            }
        )
    }

    if (clienteAEliminar != null) {
        AlertDialog(
            onDismissRequest = { clienteAEliminar = null },
            confirmButton = {
                TextButton(onClick = {
                    val cliente = clienteAEliminar ?: return@TextButton
                    viewModel.eliminarCliente(cliente.idCliente)
                    clienteAEliminar = null
                }) { Text("\uD83D\uDDD1 Eliminar") }
            },
            dismissButton = {
                TextButton(onClick = { clienteAEliminar = null }) { Text("Cancelar") }
            },
            title = { Text("Eliminar cliente") },
            text = { Text("Solo se eliminara si no tiene prestamos creados.") }
        )
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

    var busquedaCliente by remember { mutableStateOf("") }

    var clienteSeleccionado by remember { mutableStateOf<ClienteEntity?>(null) }
    var monto by remember { mutableStateOf("") }
    var interes by remember { mutableStateOf("") }
    var cuotas by remember { mutableStateOf("") }
    var intervaloDiasPersonalizado by remember { mutableStateOf("") }
    var fechaPrimeraCuota by remember { mutableStateOf(LocalDate.now()) }
    var moneda by remember { mutableStateOf(Moneda.SOLES) }
    var tipoPago by remember { mutableStateOf(TipoPago.SEMANAL) }

    var expandedMoneda by remember { mutableStateOf(false) }
    var expandedTipo by remember { mutableStateOf(false) }
    var mostrarDetallePrestamo by remember { mutableStateOf(false) }
    var prestamoDetalleId by remember { mutableStateOf<Long?>(null) }
    var filtroEstado by remember { mutableStateOf(PrestamosFiltroEstado.ACTIVOS) }
    var expandedFiltroEstado by remember { mutableStateOf(false) }
    var mostrarRegistroOk by remember { mutableStateOf(false) }

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
            Text("Prestamos", style = MaterialTheme.typography.headlineSmall)

            ClienteAutocompleteField(
                clientes = clientes,
                query = busquedaCliente,
                onQueryChange = {
                    busquedaCliente = it
                    clienteSeleccionado = null
                },
                selectedCliente = clienteSeleccionado,
                onSelectCliente = {
                    clienteSeleccionado = it
                    busquedaCliente = "${it.nombre} ${it.apellido} (${it.documentoIdentidad})"
                },
                label = "Cliente"
            )

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
                label = { Text("Interes (%)") },
                singleLine = true,
                keyboardOptions = KeyboardOptions(keyboardType = KeyboardType.Decimal),
                modifier = Modifier.fillMaxWidth()
            )

            ExposedDropdownMenuBox(expanded = expandedMoneda, onExpandedChange = { expandedMoneda = !expandedMoneda }) {
                OutlinedTextField(
                    value = when (moneda) {
                        Moneda.SOLES -> "Soles"
                        Moneda.DOLARES -> "Dolares"
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
                            text = { Text(if (opcion == Moneda.SOLES) "Soles" else "Dolares") },
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
                            text = {
                                Text(
                                    when (tipo) {
                                        TipoPago.DIARIO -> "DIARIO"
                                        TipoPago.SEMANAL -> "SEMANAL"
                                        TipoPago.QUINCENAL -> "QUINCENAL"
                                        TipoPago.MENSUAL -> "MENSUAL"
                                        TipoPago.PERSONALIZADO -> "PERSONALIZADO"
                                    }
                                )
                            },
                            onClick = {
                                tipoPago = tipo
                                expandedTipo = false
                            }
                        )
                    }
                }
            }

            if (tipoPago == TipoPago.PERSONALIZADO) {
                OutlinedTextField(
                    value = intervaloDiasPersonalizado,
                    onValueChange = { intervaloDiasPersonalizado = onlyDigits(it) },
                    label = { Text("Cada cuantos dias") },
                    singleLine = true,
                    keyboardOptions = KeyboardOptions(keyboardType = KeyboardType.Number),
                    modifier = Modifier.fillMaxWidth()
                )
            }

            val abrirCalendario = {
                DatePickerDialog(
                    context,
                    { _, year, month, dayOfMonth ->
                        fechaPrimeraCuota = LocalDate.of(year, month + 1, dayOfMonth)
                    },
                    fechaPrimeraCuota.year,
                    fechaPrimeraCuota.monthValue - 1,
                    fechaPrimeraCuota.dayOfMonth
                ).show()
            }

            OutlinedTextField(
                value = fechaPrimeraCuota.toString(),
                onValueChange = {},
                readOnly = true,
                label = { Text("Fecha primera cuota") },
                singleLine = true,
                trailingIcon = {
                    IconButton(onClick = abrirCalendario) {
                        Icon(imageVector = Icons.Outlined.DateRange, contentDescription = "Abrir calendario")
                    }
                },
                modifier = Modifier
                    .fillMaxWidth()
                    .pointerInput(fechaPrimeraCuota) {
                        detectTapGestures(onTap = { abrirCalendario() })
                    }
            )

            Button(onClick = {
                clienteSeleccionado?.let {
                    viewModel.registrarPrestamo(
                        idCliente = it.idCliente,
                        monto = monto,
                        interes = interes,
                        moneda = moneda,
                        tipoPago = tipoPago,
                        intervaloDiasPersonalizado = intervaloDiasPersonalizado,
                        cuotas = cuotas,
                        fechaPrimeraCuota = fechaPrimeraCuota.toEpochMillis()
                    ) {
                        monto = ""
                        interes = ""
                        cuotas = ""
                        intervaloDiasPersonalizado = ""
                        fechaPrimeraCuota = LocalDate.now()
                        busquedaCliente = ""
                        clienteSeleccionado = null
                        mostrarRegistroOk = true
                    }
                }
            }) { Text("Guardar prestamo") }

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
                val estadoLabel = if (prestamo.estadoPrestamo == EstadoPrestamo.ACTIVO) "\uD83D\uDFE2 ACTIVO" else "\uD83D\uDD35 PAGADO"
                val clienteLabel = "${cliente?.nombre ?: "-"} ${cliente?.apellido ?: ""}".trim()
                Column(Modifier.padding(12.dp), verticalArrangement = Arrangement.spacedBy(4.dp)) {
                    Row(
                        modifier = Modifier.fillMaxWidth(),
                        horizontalArrangement = Arrangement.SpaceBetween
                    ) {
                        Text("\uD83D\uDCB3 Prestamo #${prestamo.idPrestamo}")
                        Text(estadoLabel)
                    }
                    Text("\uD83D\uDC64 $clienteLabel")
                    Text("\uD83D\uDCB0 ${prestamo.montoTotalPrestamo.toMoney(prestamo.moneda)}")
                    Text("\uD83D\uDCCA Cuota: ${prestamo.montoCuota.toMoney(prestamo.moneda)}")
                    Spacer(Modifier.height(8.dp))
                    TextButton(onClick = { viewModel.eliminarPrestamo(prestamo.idPrestamo) }) {
                        Text("\uD83D\uDDD1 Eliminar")
                    }
                }
            }
        }
    }

    if (mostrarRegistroOk) {
        AlertDialog(
            onDismissRequest = { mostrarRegistroOk = false },
            confirmButton = {
                TextButton(onClick = { mostrarRegistroOk = false }) { Text("Aceptar") }
            },
            title = { Text("Registro") },
            text = { Text("Se realizo el registro correctamente") }
        )
    }

    if (mostrarDetallePrestamo && prestamoDetalleId != null) {
        val prestamo = prestamos.firstOrNull { it.idPrestamo == prestamoDetalleId }
        val cliente = clientes.firstOrNull { it.idCliente == prestamo?.idCliente }
        val moneda = prestamo?.moneda ?: Moneda.SOLES
        val tipoPagoDetalle = tipoPagoDetalle(prestamo?.tipoPago, cuotasDetalle)
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
            appendLine("Detalle del prestamo")
            appendLine("Cliente: ${cliente?.nombre ?: "-"} ${cliente?.apellido ?: ""}".trim())
            appendLine("Prestamo #${prestamo?.idPrestamo ?: "-"}")
            appendLine("Monto prestado: ${prestamo?.montoPrestado?.toMoney(moneda) ?: "-"}")
            appendLine("Monto total: ${prestamo?.montoTotalPrestamo?.toMoney(moneda) ?: "-"}")
            appendLine("Interes: ${prestamo?.interes ?: "-"}%")
            appendLine("Tipo pago: $tipoPagoDetalle")
            appendLine("Cuotas: ${prestamo?.cantidadCuotas ?: "-"}")
            appendLine("Fecha registro prestamo: ${prestamo?.fechaRegistro?.toDateString() ?: "-"}")
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
                    compartirTexto(context, "Detalle prestamo", detallePrestamo)
                }) {
                    Text("Compartir")
                }
            },
            dismissButton = {
                Row(horizontalArrangement = Arrangement.spacedBy(8.dp)) {
                    TextButton(onClick = {
                        runCatching {
                            createDashboardDetallePdf(context, "Detalle prestamo", detallePrestamo)
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
            title = { Text("Detalle del prestamo") },
            text = {
                LazyColumn(verticalArrangement = Arrangement.spacedBy(6.dp)) {
                    item {
                        Text("Cliente: ${cliente?.nombre ?: "-"} ${cliente?.apellido ?: ""}".trim())
                        Text("Prestamo #${prestamo?.idPrestamo ?: "-"}")
                        Text("Monto prestado: ${prestamo?.montoPrestado?.toMoney(moneda) ?: "-"}")
                        Text("Monto total: ${prestamo?.montoTotalPrestamo?.toMoney(moneda) ?: "-"}")
                        Text("Interes: ${prestamo?.interes ?: "-"}%")
                        Text("Tipo pago: $tipoPagoDetalle")
                        Text("Cuotas: ${prestamo?.cantidadCuotas ?: "-"}")
                        Text("Fecha registro prestamo: ${prestamo?.fechaRegistro?.toDateString() ?: "-"}")
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

    var busquedaCliente by remember { mutableStateOf("") }

    var idCliente by remember { mutableStateOf<Long?>(null) }
    var idPrestamo by remember { mutableStateOf<Long?>(null) }
    var idCuota by remember { mutableStateOf<Long?>(null) }
    var montoAbono by remember { mutableStateOf("") }

    var expandedPrestamo by remember { mutableStateOf(false) }
    var expandedCuota by remember { mutableStateOf(false) }
    var mostrarRegistroOk by remember { mutableStateOf(false) }

    val cuotaProxima = cuotas
        .filter { it.saldoPendiente > 0.0 }
        .minByOrNull { it.numeroCuota }
    val opcionesCuota = listOfNotNull(cuotaProxima)

    LaunchedEffect(idPrestamo, cuotaProxima?.idCuota) {
        idCuota = cuotaProxima?.idCuota
    }

    Column(
        modifier = Modifier
            .fillMaxSize()
            .padding(16.dp)
    ) {
        Text("Pagos", style = MaterialTheme.typography.headlineSmall)

        ClienteAutocompleteField(
            clientes = clientes,
            query = busquedaCliente,
            onQueryChange = {
                busquedaCliente = it
                idCliente = null
                idPrestamo = null
                idCuota = null
                viewModel.seleccionarClientePagos(null)
            },
            selectedCliente = clientes.firstOrNull { it.idCliente == idCliente },
            onSelectCliente = {
                idCliente = it.idCliente
                idPrestamo = null
                idCuota = null
                busquedaCliente = "${it.nombre} ${it.apellido} (${it.documentoIdentidad})"
                viewModel.seleccionarClientePagos(it.idCliente)
            },
            label = "Cliente"
        )

        DropdownGeneric(
            expanded = expandedPrestamo,
            onExpandedChange = { expandedPrestamo = it },
            label = "Prestamo",
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
            options = opcionesCuota,
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
        Button(onClick = {
            if (idPrestamo != null && idCuota != null) {
                viewModel.registrarPago(idPrestamo!!, idCuota!!, montoAbono) {
                    busquedaCliente = ""
                    idCliente = null
                    idPrestamo = null
                    idCuota = null
                    montoAbono = ""
                    viewModel.seleccionarClientePagos(null)
                    mostrarRegistroOk = true
                }
            }
        }) { Text("Registrar pago") }
    }

    if (mostrarRegistroOk) {
        AlertDialog(
            onDismissRequest = { mostrarRegistroOk = false },
            confirmButton = {
                TextButton(onClick = { mostrarRegistroOk = false }) { Text("Aceptar") }
            },
            title = { Text("Registro") },
            text = { Text("Se realizo el registro correctamente") }
        )
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
                    Text("Prestamos activos: ${resumen.prestamosActivos}")
                    Text("Prestamos pagados: ${resumen.prestamosPagados}")
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
                        Text("Prestamo #${cuota.idPrestamo} - Cuota ${cuota.numeroCuota}")
                        Text("Registro prestamo: ${prestamo?.fechaRegistro?.toDateString() ?: "-"}")
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

@Composable
@OptIn(ExperimentalMaterial3Api::class)
private fun ClienteAutocompleteField(
    clientes: List<ClienteEntity>,
    query: String,
    onQueryChange: (String) -> Unit,
    selectedCliente: ClienteEntity?,
    onSelectCliente: (ClienteEntity) -> Unit,
    label: String
) {
    var expanded by remember { mutableStateOf(false) }
    val clientesFiltrados = remember(clientes, query) { filtrarClientes(clientes, query).take(12) }

    Box(modifier = Modifier.fillMaxWidth()) {
        OutlinedTextField(
            value = query,
            onValueChange = {
                onQueryChange(sanitizeSingleLine(it))
                expanded = true
            },
            label = { Text(label) },
            singleLine = true,
            leadingIcon = { Icon(Icons.Outlined.Search, contentDescription = "Buscar cliente") },
            trailingIcon = {
                if (query.isNotBlank() || selectedCliente != null) {
                    IconButton(onClick = {
                        onQueryChange("")
                        expanded = false
                    }) {
                        Icon(Icons.Outlined.Close, contentDescription = "Limpiar busqueda")
                    }
                } else {
                    ExposedDropdownMenuDefaults.TrailingIcon(expanded = expanded)
                }
            },
            modifier = Modifier
                .fillMaxWidth()
        )
        DropdownMenu(
            expanded = expanded,
            onDismissRequest = { expanded = false },
            properties = PopupProperties(focusable = false)
        ) {
            if (clientesFiltrados.isEmpty()) {
                DropdownMenuItem(
                    text = { Text("Sin resultados") },
                    onClick = { expanded = false }
                )
            } else {
                clientesFiltrados.forEach { cliente ->
                    DropdownMenuItem(
                        text = { Text("\uD83D\uDC64 ${cliente.nombre} ${cliente.apellido} / \uD83D\uDCC4 ${cliente.documentoIdentidad}") },
                        onClick = {
                            onSelectCliente(cliente)
                            expanded = false
                        }
                    )
                }
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
        putExtra(Intent.EXTRA_SUBJECT, "Detalle prestamo")
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

private fun tipoPagoDetalle(tipoPago: TipoPago?, cuotas: List<com.prestamos.app.data.local.entity.CuotaEntity>): String {
    if (tipoPago == null) return "-"
    return when (tipoPago) {
        TipoPago.DIARIO -> "DIARIO"
        TipoPago.SEMANAL -> "SEMANAL"
        TipoPago.QUINCENAL -> "QUINCENAL"
        TipoPago.MENSUAL -> "MENSUAL"
        TipoPago.PERSONALIZADO -> {
            val dias = intervaloDiasEntreCuotas(cuotas)
            if (dias != null && dias > 0) "PERSONALIZADO ($dias DIAS)" else "PERSONALIZADO"
        }
    }
}

private fun intervaloDiasEntreCuotas(cuotas: List<com.prestamos.app.data.local.entity.CuotaEntity>): Long? {
    if (cuotas.size < 2) return null
    val ordenadas = cuotas.sortedBy { it.numeroCuota }
    val primera = Instant.ofEpochMilli(ordenadas[0].fechaVencimiento).atZone(ZoneId.systemDefault()).toLocalDate()
    val segunda = Instant.ofEpochMilli(ordenadas[1].fechaVencimiento).atZone(ZoneId.systemDefault()).toLocalDate()
    return ChronoUnit.DAYS.between(primera, segunda).coerceAtLeast(0L)
}
