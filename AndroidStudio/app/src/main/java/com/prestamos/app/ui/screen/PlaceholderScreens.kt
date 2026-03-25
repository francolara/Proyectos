package com.prestamos.app.ui.screen

import android.app.DatePickerDialog
import android.content.Intent
import android.widget.Toast
import androidx.compose.foundation.clickable
import androidx.compose.foundation.gestures.detectTapGestures
import androidx.compose.foundation.isSystemInDarkTheme
import androidx.compose.foundation.layout.Arrangement
import androidx.compose.foundation.layout.Box
import androidx.compose.foundation.layout.Column
import androidx.compose.foundation.layout.Row
import androidx.compose.foundation.layout.Spacer
import androidx.compose.foundation.layout.fillMaxSize
import androidx.compose.foundation.layout.fillMaxWidth
import androidx.compose.foundation.layout.height
import androidx.compose.foundation.layout.heightIn
import androidx.compose.foundation.layout.padding
import androidx.compose.foundation.lazy.LazyColumn
import androidx.compose.foundation.lazy.items
import androidx.compose.material3.AlertDialog
import androidx.compose.material3.Button
import androidx.compose.material3.Card
import androidx.compose.material3.CardDefaults
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
import androidx.compose.runtime.DisposableEffect
import androidx.compose.runtime.LaunchedEffect
import androidx.compose.runtime.getValue
import androidx.compose.runtime.mutableStateOf
import androidx.compose.runtime.remember
import androidx.compose.runtime.setValue
import androidx.compose.ui.Alignment
import androidx.compose.ui.Modifier
import androidx.compose.ui.draw.alpha
import androidx.compose.ui.graphics.Color
import androidx.compose.ui.input.pointer.pointerInput
import androidx.compose.ui.platform.LocalContext
import androidx.compose.ui.text.style.TextAlign
import androidx.compose.foundation.text.KeyboardOptions
import androidx.compose.foundation.shape.RoundedCornerShape
import androidx.compose.ui.text.input.KeyboardType
import androidx.compose.ui.unit.dp
import androidx.compose.ui.window.PopupProperties
import androidx.core.content.FileProvider
import androidx.lifecycle.Lifecycle
import androidx.lifecycle.LifecycleEventObserver
import androidx.lifecycle.compose.LocalLifecycleOwner
import androidx.lifecycle.compose.collectAsStateWithLifecycle
import com.prestamos.app.data.config.InitialSetupPreferences
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
import java.util.Locale

// Firma Codex 2026-03-21

@Composable
fun ClientesScreen(viewModel: AppViewModel) {
    val context = LocalContext.current
    val setupPrefs = remember { InitialSetupPreferences(context) }
    val visibleCurrencies = remember {
        resolveVisibleCurrencies(
            setupPrefs.getMainCurrencyCode(),
            setupPrefs.getSecondaryCurrencyCode()
        )
    }
    val clientes by viewModel.clientes.collectAsStateWithLifecycle()
    val prestamos by viewModel.prestamos.collectAsStateWithLifecycle()
    val cuotas by viewModel.cuotas.collectAsStateWithLifecycle()
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
    var clienteHistorial by remember { mutableStateOf<ClienteEntity?>(null) }

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
                Column(Modifier.padding(7.dp), verticalArrangement = Arrangement.spacedBy(2.dp)) {
                    Text(
                        "\uD83D\uDC64 ${cliente.nombre} ${cliente.apellido} / \uD83D\uDCC4 ${cliente.documentoIdentidad}",
                        style = MaterialTheme.typography.bodyMedium
                    )
                    Text("\uD83D\uDCCD ${cliente.direccion.ifBlank { "-" }}", style = MaterialTheme.typography.bodySmall)
                    Text("\uD83D\uDCDE ${cliente.telefono.ifBlank { "-" }}", style = MaterialTheme.typography.bodySmall)
                    Row(
                        modifier = Modifier.fillMaxWidth(),
                        horizontalArrangement = Arrangement.spacedBy(2.dp)
                    ) {
                        Text(
                            text = "\u270F\uFE0F Editar",
                            textAlign = TextAlign.Center,
                            style = MaterialTheme.typography.labelSmall,
                            maxLines = 1,
                            modifier = Modifier
                                .weight(1f)
                                .clickable {
                            clienteEditando = cliente
                            nombreEdit = cliente.nombre
                            apellidoEdit = cliente.apellido
                            direccionEdit = cliente.direccion
                            telefonoEdit = cliente.telefono
                                }
                                .padding(vertical = 1.dp)
                        )
                        Text(
                            text = "\uD83D\uDDD1 Eliminar",
                            textAlign = TextAlign.Center,
                            style = MaterialTheme.typography.labelSmall,
                            maxLines = 1,
                            modifier = Modifier
                                .weight(1f)
                                .clickable { clienteAEliminar = cliente }
                                .padding(vertical = 1.dp)
                        )
                        Text(
                            text = "\uD83D\uDCDC Historial",
                            textAlign = TextAlign.Center,
                            style = MaterialTheme.typography.labelSmall,
                            maxLines = 1,
                            modifier = Modifier
                                .weight(1f)
                                .clickable { clienteHistorial = cliente }
                                .padding(vertical = 1.dp)
                        )
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

    if (clienteHistorial != null) {
        val cliente = clienteHistorial ?: return
        val prestamosCliente = prestamos
            .filter { it.idCliente == cliente.idCliente && (it.estadoPrestamo == EstadoPrestamo.ACTIVO || it.estadoPrestamo == EstadoPrestamo.PAGADO) }
            .sortedByDescending { it.fechaRegistro }
        val moraPendientePorPrestamo = prestamosCliente.associate { prestamo ->
            prestamo.idPrestamo to cuotas
                .filter { it.idPrestamo == prestamo.idPrestamo }
                .sumOf { it.moraPendiente }
                .coerceAtLeast(0.0)
        }
        val saldoPendienteConInteresPorPrestamo = prestamosCliente.associate { prestamo ->
            prestamo.idPrestamo to cuotas
                .filter { it.idPrestamo == prestamo.idPrestamo }
                .sumOf { it.saldoPendiente + it.moraPendiente }
                .coerceAtLeast(0.0)
        }

        val totalCapitalPorMoneda = prestamosCliente.groupBy { it.moneda }.mapValues { (_, v) -> v.sumOf { it.montoPrestado } }
        val totalCapitalConInteresPorMoneda = prestamosCliente.groupBy { it.moneda }.mapValues { (_, v) ->
            v.sumOf { prestamo ->
                prestamo.montoTotalPrestamo + (moraPendientePorPrestamo[prestamo.idPrestamo] ?: 0.0)
            }
        }
        val totalPendienteConInteresPorMoneda = prestamosCliente.groupBy { it.moneda }.mapValues { (_, v) ->
            v.sumOf { prestamo ->
                saldoPendienteConInteresPorPrestamo[prestamo.idPrestamo] ?: 0.0
            }
        }
        val monedasConPrestamos = prestamosCliente.map { it.moneda }
        val monedasResumen = (visibleCurrencies + monedasConPrestamos)
            .distinct()
            .ifEmpty { listOf(Moneda.SOLES) }

        val historialTexto = buildString {
            appendLine("Historial de prestamos")
            appendLine("Cliente: ${cliente.nombre} ${cliente.apellido}".trim())
            appendLine("Documento: ${cliente.documentoIdentidad}")
            if (prestamosCliente.isEmpty()) {
                appendLine("No hay prestamos activos o pagados para este cliente.")
            } else {
                appendLine()
                appendLine("Resumen general (por moneda)")
                monedasResumen.forEach { moneda ->
                    appendLine("Moneda: ${moneda.displayName}")
                    appendLine("  Capital: ${(totalCapitalPorMoneda[moneda] ?: 0.0).toMoney(moneda)}")
                    appendLine("  Capital + intereses: ${(totalCapitalConInteresPorMoneda[moneda] ?: 0.0).toMoney(moneda)}")
                    appendLine("  Saldo pendiente + intereses: ${(totalPendienteConInteresPorMoneda[moneda] ?: 0.0).toMoney(moneda)}")
                    appendLine("  Pendiente: ${(totalPendienteConInteresPorMoneda[moneda] ?: 0.0).toMoney(moneda)}")
                }
                appendLine()

                prestamosCliente.forEach { prestamo ->
                    val saldoPendienteConInteres = saldoPendienteConInteresPorPrestamo[prestamo.idPrestamo] ?: 0.0
                    val moraPendientePrestamo = moraPendientePorPrestamo[prestamo.idPrestamo] ?: 0.0
                    val estadoTexto = if (prestamo.estadoPrestamo == EstadoPrestamo.ACTIVO) "ACTIVO" else "PAGADO"
                    appendLine("Prestamo #${prestamo.idPrestamo} - $estadoTexto")
                    appendLine("  Fecha: ${prestamo.fechaRegistro.toDateString()}")
                    appendLine("  Capital: ${prestamo.montoPrestado.toMoney(prestamo.moneda)}")
                    appendLine("  Capital + intereses: ${(prestamo.montoTotalPrestamo + moraPendientePrestamo).toMoney(prestamo.moneda)}")
                    appendLine("  Pendiente + intereses: ${saldoPendienteConInteres.toMoney(prestamo.moneda)}")
                }
            }
        }

        AlertDialog(
            onDismissRequest = { clienteHistorial = null },
            confirmButton = {
                TextButton(onClick = {
                    compartirTexto(
                        context = context,
                        titulo = "Historial de prestamos - ${cliente.nombre} ${cliente.apellido}",
                        detalle = historialTexto
                    )
                }) {
                    Text("Compartir")
                }
            },
            dismissButton = {
                Row(horizontalArrangement = Arrangement.spacedBy(8.dp)) {
                    TextButton(onClick = {
                        runCatching {
                            createDashboardDetallePdf(
                                context,
                                "Historial de prestamos - ${cliente.nombre} ${cliente.apellido}",
                                historialTexto
                            )
                        }.onSuccess { file ->
                            compartirArchivo(context, file, "application/pdf")
                        }.onFailure {
                            Toast.makeText(context, "No se pudo exportar PDF", Toast.LENGTH_SHORT).show()
                        }
                    }) {
                        Text("PDF")
                    }
                    TextButton(onClick = { clienteHistorial = null }) {
                        Text("Cerrar")
                    }
                }
            },
            title = { Text("Historial de prestamos") },
            text = {
                LazyColumn(
                    verticalArrangement = Arrangement.spacedBy(8.dp),
                    modifier = Modifier
                        .fillMaxWidth()
                        .heightIn(max = 440.dp)
                ) {
                    if (prestamosCliente.isEmpty()) {
                        item {
                            Text("No hay prestamos activos o pagados para este cliente.")
                        }
                    } else {
                        item {
                            Text("\uD83D\uDCCA Resumen general", style = MaterialTheme.typography.labelLarge)
                        }
                        items(monedasResumen) { moneda ->
                            Card(modifier = Modifier.fillMaxWidth()) {
                                Column(
                                    modifier = Modifier.padding(10.dp),
                                    verticalArrangement = Arrangement.spacedBy(3.dp)
                                ) {
                                    Text("\uD83D\uDCB1 ${moneda.displayName}", style = MaterialTheme.typography.titleSmall)
                                    Text("\uD83D\uDCB0 Capital: ${(totalCapitalPorMoneda[moneda] ?: 0.0).toMoney(moneda)}")
                                    Text("\uD83D\uDCB0 Capital + intereses: ${(totalCapitalConInteresPorMoneda[moneda] ?: 0.0).toMoney(moneda)}")
                                    Text("\uD83D\uDCB5 Saldo pendiente + intereses: ${(totalPendienteConInteresPorMoneda[moneda] ?: 0.0).toMoney(moneda)}")
                                    Text("\uD83E\uDDFE Pendiente: ${(totalPendienteConInteresPorMoneda[moneda] ?: 0.0).toMoney(moneda)}")
                                }
                            }
                        }
                        item {
                            Text("\uD83D\uDCC4 Lista", style = MaterialTheme.typography.labelLarge)
                        }
                        items(prestamosCliente, key = { it.idPrestamo }) { prestamo ->
                            val saldoPendienteConInteres = saldoPendienteConInteresPorPrestamo[prestamo.idPrestamo] ?: 0.0
                            val moraPendientePrestamo = moraPendientePorPrestamo[prestamo.idPrestamo] ?: 0.0
                            val activo = prestamo.estadoPrestamo == EstadoPrestamo.ACTIVO
                            val estadoTexto = if (activo) "\uD83D\uDFE1 ACTIVO" else "\uD83D\uDFE2 PAGADO"
                            val estadoColor = if (activo) MaterialTheme.colorScheme.tertiary else MaterialTheme.colorScheme.primary

                            Card(modifier = Modifier.fillMaxWidth()) {
                                Column(
                                    modifier = Modifier.padding(10.dp),
                                    verticalArrangement = Arrangement.spacedBy(3.dp)
                                ) {
                                    Text("\uD83D\uDCC4 Prestamo #${prestamo.idPrestamo}        $estadoTexto", style = MaterialTheme.typography.titleSmall, color = estadoColor)
                                    Text("\uD83D\uDCC5 ${prestamo.fechaRegistro.toDateString()}")
                                    Text("\uD83D\uDCB0 ${prestamo.montoPrestado.toMoney(prestamo.moneda)}  ${prestamo.moneda.displayName}")
                                    Text("\uD83D\uDCB5 ${(prestamo.montoTotalPrestamo + moraPendientePrestamo).toMoney(prestamo.moneda)}")
                                    Text("${if (activo) "\uD83D\uDFE1" else "\u2714\uFE0F"} ${saldoPendienteConInteres.toMoney(prestamo.moneda)}")
                                }
                            }
                        }
                    }
                }
            }
        )
    }
}

private fun Map<Moneda, Double>.toTotalsText(visibleCurrencies: List<Moneda>): String {
    val ordered = visibleCurrencies.ifEmpty { listOf(Moneda.SOLES) }
    return ordered.joinToString("\n") { moneda ->
        val total = this[moneda] ?: 0.0
        val label = "Totales en ${moneda.displayName}"
        "$label: ${total.toMoney(moneda)}"
    }
}

private fun resolveVisibleCurrencies(mainCode: String?, secondaryCode: String?): List<Moneda> {
    val mapped = listOfNotNull(mainCode.toMonedaOrNull(), secondaryCode.toMonedaOrNull())
        .distinct()
    return if (mapped.isEmpty()) listOf(Moneda.SOLES) else mapped
}

private fun String?.toMonedaOrNull(): Moneda? = Moneda.fromCode(this)

private fun Moneda.toDisplayName(): String = "${symbol} - $displayName"

private enum class PrestamosFiltroEstado(val estado: EstadoPrestamo, val label: String) {
    ACTIVOS(EstadoPrestamo.ACTIVO, "ACTIVO"),
    PAGADOS(EstadoPrestamo.PAGADO, "PAGADO")
}

@OptIn(ExperimentalMaterial3Api::class)
@Composable
fun PrestamosScreen(viewModel: AppViewModel) {
    val context = LocalContext.current
    val setupPrefs = remember { InitialSetupPreferences(context) }
    var monedasDisponibles by remember {
        mutableStateOf(
            resolveVisibleCurrencies(
                setupPrefs.getMainCurrencyCode(),
                setupPrefs.getSecondaryCurrencyCode()
            )
        )
    }
    var tiposPagoDisponibles by remember { mutableStateOf(setupPrefs.getAllowedPaymentTypes().toList()) }
    var interesPorDefecto by remember { mutableStateOf(setupPrefs.getDefaultInterest()) }
    val lifecycleOwner = LocalLifecycleOwner.current
    val refreshMonedas: () -> Unit = {
        monedasDisponibles = resolveVisibleCurrencies(
            setupPrefs.getMainCurrencyCode(),
            setupPrefs.getSecondaryCurrencyCode()
        )
        tiposPagoDisponibles = setupPrefs.getAllowedPaymentTypes().toList()
        interesPorDefecto = setupPrefs.getDefaultInterest()
    }
    DisposableEffect(lifecycleOwner) {
        val observer = LifecycleEventObserver { _, event ->
            if (event == Lifecycle.Event.ON_RESUME) {
                refreshMonedas()
            }
        }
        lifecycleOwner.lifecycle.addObserver(observer)
        onDispose { lifecycleOwner.lifecycle.removeObserver(observer) }
    }

    val clientes by viewModel.clientes.collectAsStateWithLifecycle()
    val prestamos by viewModel.prestamos.collectAsStateWithLifecycle()

    var busquedaCliente by remember { mutableStateOf("") }

    var clienteSeleccionado by remember { mutableStateOf<ClienteEntity?>(null) }
    var monto by remember { mutableStateOf("") }
    var interes by remember { mutableStateOf(interesPorDefecto) }
    var cuotas by remember { mutableStateOf("") }
    var intervaloDiasPersonalizado by remember { mutableStateOf("") }
    var fechaPrimeraCuota by remember { mutableStateOf(LocalDate.now()) }
    var moneda by remember { mutableStateOf(monedasDisponibles.firstOrNull() ?: Moneda.SOLES) }
    var tipoPago by remember { mutableStateOf(tiposPagoDisponibles.firstOrNull() ?: TipoPago.SEMANAL) }

    LaunchedEffect(monedasDisponibles) {
        if (moneda !in monedasDisponibles) {
            moneda = monedasDisponibles.firstOrNull() ?: Moneda.SOLES
        }
    }
    LaunchedEffect(tiposPagoDisponibles) {
        if (tiposPagoDisponibles.isNotEmpty() && tipoPago !in tiposPagoDisponibles) {
            tipoPago = tiposPagoDisponibles.first()
        }
        if (interes.isBlank()) {
            interes = interesPorDefecto
        }
    }
    LaunchedEffect(interesPorDefecto) {
        if (interes.isBlank()) {
            interes = interesPorDefecto
        }
    }

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
    val pagos by viewModel.pagos.collectAsStateWithLifecycle()
    val tiposCobro by viewModel.tiposCobro.collectAsStateWithLifecycle()

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
                textStyle = MaterialTheme.typography.bodyLarge.copy(textAlign = TextAlign.End),
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
                    value = moneda.toDisplayName(),
                    onValueChange = {},
                    readOnly = true,
                    label = { Text("Moneda") },
                    trailingIcon = { ExposedDropdownMenuDefaults.TrailingIcon(expanded = expandedMoneda) },
                    modifier = Modifier.menuAnchor().fillMaxWidth()
                )
                DropdownMenu(expanded = expandedMoneda, onDismissRequest = { expandedMoneda = false }) {
                    monedasDisponibles.forEach { opcion ->
                        DropdownMenuItem(
                            text = { Text(opcion.toDisplayName()) },
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

            if (tiposPagoDisponibles.size <= 1) {
                OutlinedTextField(
                    value = tipoPago.toUiLabel(),
                    onValueChange = {},
                    readOnly = true,
                    label = { Text("Tipo de pago") },
                    modifier = Modifier.fillMaxWidth()
                )
            } else {
                ExposedDropdownMenuBox(expanded = expandedTipo, onExpandedChange = { expandedTipo = !expandedTipo }) {
                    OutlinedTextField(
                        value = tipoPago.toUiLabel(),
                        onValueChange = {},
                        readOnly = true,
                        label = { Text("Tipo de pago") },
                        trailingIcon = { ExposedDropdownMenuDefaults.TrailingIcon(expanded = expandedTipo) },
                        modifier = Modifier.menuAnchor().fillMaxWidth()
                    )
                    DropdownMenu(expanded = expandedTipo, onDismissRequest = { expandedTipo = false }) {
                        tiposPagoDisponibles.forEach { tipo ->
                            DropdownMenuItem(
                                text = { Text(tipo.toUiLabel()) },
                                onClick = {
                                    tipoPago = tipo
                                    expandedTipo = false
                                }
                            )
                        }
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
                        interes = interes.ifBlank { "0" },
                        moneda = moneda,
                        tipoPago = tipoPago,
                        intervaloDiasPersonalizado = intervaloDiasPersonalizado,
                        cuotas = cuotas,
                        fechaPrimeraCuota = fechaPrimeraCuota.toEpochMillis()
                    ) {
                        monto = ""
                        interes = interesPorDefecto
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
                Column(Modifier.padding(7.dp), verticalArrangement = Arrangement.spacedBy(2.dp)) {
                    Row(
                        modifier = Modifier.fillMaxWidth(),
                        horizontalArrangement = Arrangement.SpaceBetween
                    ) {
                        Text("\uD83D\uDCC4 Prestamo #${prestamo.idPrestamo}", style = MaterialTheme.typography.bodyMedium)
                        Text(estadoLabel, style = MaterialTheme.typography.labelSmall)
                    }
                    Text("\uD83D\uDC64 $clienteLabel", style = MaterialTheme.typography.bodySmall)
                    Text(
                        "\uD83D\uDCB0 ${prestamo.montoTotalPrestamo.toMoney(prestamo.moneda)} ${prestamo.moneda.displayName}",
                        style = MaterialTheme.typography.bodySmall
                    )
                    Text(
                        "\uD83D\uDCCC Cuota: ${prestamo.montoCuota.toMoney(prestamo.moneda)} ${prestamo.moneda.displayName}",
                        style = MaterialTheme.typography.bodySmall
                    )
                    Text(
                        text = "\uD83D\uDDD1 Eliminar",
                        style = MaterialTheme.typography.labelSmall,
                        modifier = Modifier
                            .align(Alignment.End)
                            .clickable { viewModel.eliminarPrestamo(prestamo.idPrestamo) }
                            .padding(top = 1.dp, bottom = 0.dp)
                    )
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
        val pagosPrestamo = pagos.filter { it.idPrestamo == prestamoDetalleId }
        val moraCobradaPorCuota = pagosPrestamo
            .groupBy { it.idCuota }
            .mapValues { (_, lista) -> lista.sumOf { it.moraCobrada } }
        val tipoCobroPorCuota = pagos
            .filter { it.idPrestamo == prestamoDetalleId && it.idTipoCobro != null }
            .groupBy { it.idCuota }
            .mapValues { (_, lista) -> lista.maxByOrNull { it.fechaPago }?.idTipoCobro }
        val tipoCobroNombreById = tiposCobro.associateBy { it.idTipoCobro }
        val cronograma = if (cuotasDetalle.isEmpty()) {
            "Sin cuotas registradas"
        } else {
            cuotasDetalle.joinToString("\n") { cuota ->
                val moraCobrada = moraCobradaPorCuota[cuota.idCuota] ?: 0.0
                val moraPendiente = cuota.moraPendiente
                val moraTotal = moraCobrada + moraPendiente
                val totalCuotaConMora = cuota.montoCuota + moraTotal
                val pendienteSinMora = (cuota.montoCuota - cuota.montoPagado).coerceAtLeast(0.0)
                val pendienteConMora = pendienteSinMora + moraPendiente
                val tipoCobro = tipoCobroPorCuota[cuota.idCuota]?.let { tipoCobroNombreById[it]?.nombre }
                val tipoCobroTexto = if (cuota.estadoCuota.name == "PAGADO" && !tipoCobro.isNullOrBlank()) {
                    " | tipo cobro $tipoCobro"
                } else {
                    ""
                }
                "Cuota ${cuota.numeroCuota}: vence ${cuota.fechaVencimiento.toDateString()} | " +
                    "monto ${cuota.montoCuota.toMoney(moneda)} | " +
                    "mora ${moraTotal.toMoney(moneda)} | " +
                    "total ${totalCuotaConMora.toMoney(moneda)} | " +
                    "pendiente ${pendienteConMora.toMoney(moneda)}$tipoCobroTexto"
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
                LazyColumn(
                    verticalArrangement = Arrangement.spacedBy(8.dp),
                    modifier = Modifier.heightIn(max = 430.dp)
                ) {
                    item {
                        Card(modifier = Modifier.fillMaxWidth()) {
                            Column(
                                modifier = Modifier
                                    .fillMaxWidth()
                                    .padding(8.dp),
                                verticalArrangement = Arrangement.spacedBy(3.dp)
                            ) {
                                Text("\uD83D\uDC64 Cliente", style = MaterialTheme.typography.labelLarge)
                                Text(
                                    "${cliente?.nombre ?: "-"} ${cliente?.apellido ?: ""}".trim(),
                                    style = MaterialTheme.typography.bodyLarge
                                )
                                Text("\uD83D\uDCC4 Prestamo #${prestamo?.idPrestamo ?: "-"}", style = MaterialTheme.typography.labelSmall)
                            }
                        }
                    }
                    item {
                        Card(modifier = Modifier.fillMaxWidth()) {
                            Column(
                                modifier = Modifier
                                    .fillMaxWidth()
                                    .padding(8.dp),
                                verticalArrangement = Arrangement.spacedBy(3.dp)
                            ) {
                                Text("\uD83D\uDCB0 Resumen", style = MaterialTheme.typography.labelLarge)
                                Text(
                                    "\uD83D\uDCB5 Monto prestado: ${prestamo?.montoPrestado?.toMoney(moneda) ?: "-"}",
                                    style = MaterialTheme.typography.bodySmall
                                )
                                Text(
                                    "\uD83D\uDCB8 Total a pagar: ${prestamo?.montoTotalPrestamo?.toMoney(moneda) ?: "-"}",
                                    style = MaterialTheme.typography.bodySmall
                                )
                                Text(
                                    "\uD83E\uDDFE Deuda pendiente: ${totalDeudaPendiente.toMoney(moneda)} ${moneda.displayName}",
                                    style = MaterialTheme.typography.bodySmall
                                )
                            }
                        }
                    }
                    item {
                        Card(modifier = Modifier.fillMaxWidth()) {
                            Column(
                                modifier = Modifier
                                    .fillMaxWidth()
                                    .padding(8.dp),
                                verticalArrangement = Arrangement.spacedBy(3.dp)
                            ) {
                                Text("\u2699\uFE0F Condiciones", style = MaterialTheme.typography.labelLarge)
                                Text("\uD83D\uDCCA Interes: ${prestamo?.interes ?: "-"}%", style = MaterialTheme.typography.bodySmall)
                                Text("\uD83D\uDD01 Frecuencia: $tipoPagoDetalle", style = MaterialTheme.typography.bodySmall)
                                Text("\uD83D\uDCCC Cuotas: ${prestamo?.cantidadCuotas ?: "-"}", style = MaterialTheme.typography.bodySmall)
                                Text("\uD83D\uDCC5 Fecha de registro: ${prestamo?.fechaRegistro?.toDateString() ?: "-"}", style = MaterialTheme.typography.bodySmall)
                            }
                        }
                    }
                    item {
                        Text("\uD83D\uDCCC Cronograma de cuotas", style = MaterialTheme.typography.titleSmall)
                    }
                    items(cuotasDetalle) { cuota ->
                        val estadoTexto = when (cuota.estadoCuota.name) {
                            "PAGADO" -> "\uD83D\uDFE2 Pagado"
                            "PARCIAL" -> "\uD83D\uDFE0 Parcial"
                            "VENCIDO" -> "\uD83D\uDD34 Vencido"
                            else -> "\uD83D\uDFE1 Pendiente"
                        }
                        Card(
                            modifier = Modifier.fillMaxWidth(),
                            colors = androidx.compose.material3.CardDefaults.cardColors(
                                containerColor = MaterialTheme.colorScheme.secondaryContainer.copy(alpha = 0.55f)
                            )
                        ) {
                            Column(
                                modifier = Modifier
                                    .fillMaxWidth()
                                    .padding(8.dp),
                                verticalArrangement = Arrangement.spacedBy(2.dp)
                            ) {
                                val moraCobrada = moraCobradaPorCuota[cuota.idCuota] ?: 0.0
                                val moraPendiente = cuota.moraPendiente
                                val moraTotal = moraCobrada + moraPendiente
                                val totalCuotaConMora = cuota.montoCuota + moraTotal
                                val pendienteSinMora = (cuota.montoCuota - cuota.montoPagado).coerceAtLeast(0.0)
                                val pendienteConMora = pendienteSinMora + moraPendiente

                                Text("\uD83D\uDCCC Cuota ${cuota.numeroCuota}", style = MaterialTheme.typography.labelMedium)
                                Text("\uD83D\uDCC5 ${cuota.fechaVencimiento.toDateString()}", style = MaterialTheme.typography.labelSmall)
                                Text("\uD83D\uDCB0 Cuota base: ${cuota.montoCuota.toMoney(moneda)}", style = MaterialTheme.typography.labelSmall)
                                if (moraTotal > 0.0) {
                                    Text(
                                        "\uD83D\uDCB8 Mora: ${moraTotal.toMoney(moneda)} (cobrada ${moraCobrada.toMoney(moneda)} / pendiente ${moraPendiente.toMoney(moneda)})",
                                        style = MaterialTheme.typography.labelSmall
                                    )
                                }
                                Text("\uD83D\uDCB5 Total cuota + mora: ${totalCuotaConMora.toMoney(moneda)}", style = MaterialTheme.typography.labelSmall)
                                Text("\uD83E\uDDFE Pendiente: ${pendienteConMora.toMoney(moneda)}", style = MaterialTheme.typography.labelSmall)
                                if (cuota.estadoCuota.name == "PAGADO") {
                                    Text("\u2705 Cobrado: ${(cuota.montoPagado + moraCobrada).toMoney(moneda)}", style = MaterialTheme.typography.labelSmall)
                                    val tipoCobro = tipoCobroPorCuota[cuota.idCuota]?.let { tipoCobroNombreById[it]?.nombre }
                                    if (!tipoCobro.isNullOrBlank()) {
                                        Text("\uD83D\uDCB8 $tipoCobro", style = MaterialTheme.typography.labelSmall)
                                    }
                                }
                                Text(estadoTexto, style = MaterialTheme.typography.labelSmall)
                            }
                        }
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
    val prestamosCliente by viewModel.prestamosClientePagos.collectAsStateWithLifecycle()
    val cuotas by viewModel.cuotasPrestamoPagos.collectAsStateWithLifecycle()
    val prestamosTodos by viewModel.prestamos.collectAsStateWithLifecycle()
    val cuotasTodas by viewModel.cuotas.collectAsStateWithLifecycle()
    val pagos by viewModel.pagos.collectAsStateWithLifecycle()
    val tiposCobro by viewModel.tiposCobro.collectAsStateWithLifecycle()

    var busquedaCliente by remember { mutableStateOf("") }

    var idCliente by remember { mutableStateOf<Long?>(null) }
    var idPrestamo by remember { mutableStateOf<Long?>(null) }
    var idCuota by remember { mutableStateOf<Long?>(null) }
    var idTipoCobro by remember { mutableStateOf<Long?>(null) }
    var montoAbono by remember { mutableStateOf("") }

    var expandedPrestamo by remember { mutableStateOf(false) }
    var expandedCuota by remember { mutableStateOf(false) }
    var expandedTipoCobro by remember { mutableStateOf(false) }
    var mostrarRegistroOk by remember { mutableStateOf(false) }
    var pagoAEliminar by remember { mutableStateOf<PagoListadoItem?>(null) }

    val cuotaProxima = cuotas
        .filter { it.saldoPendiente > 0.0 }
        .minByOrNull { it.numeroCuota }
    val opcionesCuota = listOfNotNull(cuotaProxima)
    val clienteById = remember(clientes) { clientes.associateBy { it.idCliente } }
    val prestamoById = remember(prestamosTodos) { prestamosTodos.associateBy { it.idPrestamo } }
    val cuotaById = remember(cuotasTodas) { cuotasTodas.associateBy { it.idCuota } }
    val tipoCobroById = remember(tiposCobro) { tiposCobro.associateBy { it.idTipoCobro } }
    val ultimoPagoIdPorPrestamo = remember(pagos) {
        pagos.groupBy { it.idPrestamo }.mapValues { (_, lista) -> lista.firstOrNull()?.idPago }
    }
    val pagosListado = remember(pagos, prestamoById, clienteById, cuotaById, tipoCobroById) {
        pagos.mapNotNull { pago ->
            val prestamo = prestamoById[pago.idPrestamo] ?: return@mapNotNull null
            val cliente = clienteById[prestamo.idCliente]
            val cuota = cuotaById[pago.idCuota]
            val tipoCobroNombre = pago.idTipoCobro?.let { tipoCobroById[it]?.nombre } ?: "-"
            PagoListadoItem(
                idPago = pago.idPago,
                idPrestamo = pago.idPrestamo,
                numeroCuota = cuota?.numeroCuota ?: 0,
                clienteNombre = cliente?.let { "${it.nombre} ${it.apellido}".trim() } ?: "Cliente no disponible",
                montoAbono = pago.montoAbono,
                fechaPago = pago.fechaPago,
                moneda = prestamo.moneda,
                cuotaPagada = (cuota?.saldoPendiente ?: 0.0) <= 0.0,
                tipoCobro = tipoCobroNombre
            )
        }
    }

    LaunchedEffect(idPrestamo, cuotaProxima?.idCuota) {
        idCuota = cuotaProxima?.idCuota
    }
    LaunchedEffect(tiposCobro) {
        if (tiposCobro.size == 1 && idTipoCobro == null) {
            idTipoCobro = tiposCobro.first().idTipoCobro
        }
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
            selected = prestamosCliente.firstOrNull { it.idPrestamo == idPrestamo }
                ?.let { "#${it.idPrestamo} - saldo ${it.montoTotalPrestamo.toMoney(it.moneda)}" } ?: "",
            options = prestamosCliente,
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

        DropdownGeneric(
            expanded = expandedTipoCobro,
            onExpandedChange = { expandedTipoCobro = it },
            label = "Tipo de cobro",
            selected = tiposCobro.firstOrNull { it.idTipoCobro == idTipoCobro }?.nombre ?: "",
            options = tiposCobro,
            optionText = { it.nombre },
            onSelect = { idTipoCobro = it.idTipoCobro }
        )

        OutlinedTextField(
            value = montoAbono,
            onValueChange = { montoAbono = onlyDecimal(it) },
            label = { Text("Monto abonado") },
            singleLine = true,
            textStyle = MaterialTheme.typography.bodyLarge.copy(textAlign = TextAlign.End),
            keyboardOptions = KeyboardOptions(keyboardType = KeyboardType.Decimal),
            modifier = Modifier.fillMaxWidth()
        )
        Button(onClick = {
            if (idPrestamo != null && idCuota != null) {
                viewModel.registrarPago(idPrestamo!!, idCuota!!, idTipoCobro, montoAbono) {
                    busquedaCliente = ""
                    idCliente = null
                    idPrestamo = null
                    idCuota = null
                    idTipoCobro = if (tiposCobro.size == 1) tiposCobro.first().idTipoCobro else null
                    montoAbono = ""
                    viewModel.seleccionarClientePagos(null)
                    mostrarRegistroOk = true
                }
            }
        }) { Text("Registrar pago") }

        Spacer(modifier = Modifier.height(12.dp))
        Text("Listado", style = MaterialTheme.typography.titleMedium)
        Spacer(modifier = Modifier.height(8.dp))
        if (pagosListado.isEmpty()) {
            Text("No hay pagos registrados.")
        } else {
            LazyColumn(
                modifier = Modifier
                    .fillMaxWidth()
                    .weight(1f),
                verticalArrangement = Arrangement.spacedBy(4.dp)
            ) {
                items(pagosListado, key = { it.idPago }) { pago ->
                    val puedeEliminar = ultimoPagoIdPorPrestamo[pago.idPrestamo] == pago.idPago
                    Card(
                        modifier = Modifier.fillMaxWidth(),
                        shape = RoundedCornerShape(12.dp)
                    ) {
                        Column(
                            modifier = Modifier
                                .fillMaxWidth()
                                .padding(7.dp),
                            verticalArrangement = Arrangement.spacedBy(2.dp)
                        ) {
                            Row(
                                modifier = Modifier.fillMaxWidth(),
                                horizontalArrangement = Arrangement.SpaceBetween
                            ) {
                                Text(
                                    text = "\uD83D\uDC64 ${pago.clienteNombre}",
                                    style = MaterialTheme.typography.bodyMedium
                                )
                                Text(
                                    text = "\uD83D\uDCC5 ${pago.fechaPago.toDateString()}",
                                    style = MaterialTheme.typography.bodySmall,
                                    color = MaterialTheme.colorScheme.onSurfaceVariant
                                )
                            }
                            Text(
                                text = "\uD83D\uDCC4 Prestamo #${pago.idPrestamo}",
                                style = MaterialTheme.typography.bodySmall
                            )
                            Text(
                                text = "\uD83D\uDCCC Cuota ${pago.numeroCuota}",
                                style = MaterialTheme.typography.bodySmall
                            )
                            Text(
                                text = "\uD83D\uDCB8 ${pago.tipoCobro}",
                                style = MaterialTheme.typography.bodySmall
                            )
                            Row(
                                modifier = Modifier.fillMaxWidth(),
                                horizontalArrangement = Arrangement.SpaceBetween
                            ) {
                                Text(
                                    text = "\uD83D\uDCB0 ${pago.montoAbono.toMoney(pago.moneda)}",
                                    style = MaterialTheme.typography.bodySmall
                                )
                                Text(
                                    text = if (pago.cuotaPagada) "\u2705 Pagado" else "\uD83D\uDFE1 Parcial",
                                    style = MaterialTheme.typography.bodySmall,
                                    color = if (pago.cuotaPagada) Color(0xFF2E7D32) else Color(0xFFF57F17)
                                )
                            }
                            Text(
                                text = "\uD83D\uDDD1 Eliminar",
                                style = MaterialTheme.typography.labelSmall,
                                color = if (puedeEliminar) MaterialTheme.colorScheme.error else MaterialTheme.colorScheme.onSurfaceVariant,
                                modifier = Modifier
                                    .align(Alignment.End)
                                    .alpha(if (puedeEliminar) 1f else 0.45f)
                                    .clickable(enabled = puedeEliminar) { pagoAEliminar = pago }
                                    .padding(top = 1.dp, bottom = 0.dp)
                            )
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

    if (pagoAEliminar != null) {
        val pago = pagoAEliminar ?: return
        AlertDialog(
            onDismissRequest = { pagoAEliminar = null },
            confirmButton = {
                TextButton(onClick = {
                    viewModel.eliminarPago(pago.idPago)
                    pagoAEliminar = null
                }) {
                    Text("Eliminar")
                }
            },
            dismissButton = {
                TextButton(onClick = { pagoAEliminar = null }) { Text("Cancelar") }
            },
            title = { Text("Eliminar pago") },
            text = { Text("Solo se puede eliminar el ultimo pago del prestamo.\n\nSe eliminara el pago #${pago.idPago} (${pago.tipoCobro}).") }
        )
    }
}

private data class PagoListadoItem(
    val idPago: Long,
    val idPrestamo: Long,
    val numeroCuota: Int,
    val clienteNombre: String,
    val montoAbono: Double,
    val fechaPago: Long,
    val moneda: Moneda,
    val cuotaPagada: Boolean,
    val tipoCobro: String
)

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
                        Text("\uD83D\uDCC4 Prestamo #${cuota.idPrestamo} - \uD83D\uDCCC Cuota ${cuota.numeroCuota}")
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

private fun TipoPago.toUiLabel(): String = when (this) {
    TipoPago.DIARIO -> "DIARIO"
    TipoPago.SEMANAL -> "SEMANAL"
    TipoPago.QUINCENAL -> "QUINCENAL"
    TipoPago.MENSUAL -> "MENSUAL"
    TipoPago.PERSONALIZADO -> "PERSONALIZADO"
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
