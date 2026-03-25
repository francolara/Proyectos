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
import com.prestamos.app.data.local.entity.PagoEntity
import com.prestamos.app.data.local.entity.TipoPago
import com.prestamos.app.data.local.entity.TipoCobroEntity
import com.prestamos.app.ui.screen.export.ReportPdfPayload
import com.prestamos.app.ui.screen.export.ReportPdfSection
import com.prestamos.app.ui.screen.export.ReportTable
import com.prestamos.app.ui.screen.export.ReportTableColumn
import com.prestamos.app.ui.screen.export.TableAlign
import com.prestamos.app.ui.screen.export.createReportesPdf
import com.prestamos.app.ui.screen.export.createDashboardDetallePdf
import com.prestamos.app.ui.viewmodel.AppViewModel
import com.prestamos.app.util.toDateString
import com.prestamos.app.util.toEpochMillis
import com.prestamos.app.util.toMoney
import java.io.File
import java.text.SimpleDateFormat
import java.time.Instant
import java.time.LocalDate
import java.time.ZoneId
import java.time.temporal.ChronoUnit
import java.util.Date
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

        if (clientes.isEmpty()) {
            item { Text("No hay clientes registrados.") }
        } else {
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

        if (prestamosFiltrados.isEmpty()) {
            item { Text("No hay prestamos registrados.") }
        } else {
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
                val moraTexto = if (moraTotal > 0.0) {
                    " | mora ${moraTotal.toMoney(moneda)}"
                } else {
                    ""
                }
                val totalCuotaConMoraTexto = if (moraTotal > 0.0) {
                    " | total cuota + mora ${totalCuotaConMora.toMoney(moneda)}"
                } else {
                    ""
                }
                "Cuota ${cuota.numeroCuota}: vence ${cuota.fechaVencimiento.toDateString()} | " +
                    "monto ${cuota.montoCuota.toMoney(moneda)} | " +
                    "pendiente ${pendienteConMora.toMoney(moneda)}$moraTexto$totalCuotaConMoraTexto$tipoCobroTexto"
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
                                        "\uD83D\uDCB8 Mora: ${moraTotal.toMoney(moneda)}",
                                        style = MaterialTheme.typography.labelSmall
                                    )
                                }
                                if (moraTotal > 0.0) {
                                    Text("\uD83D\uDCB5 Total cuota + mora: ${totalCuotaConMora.toMoney(moneda)}", style = MaterialTheme.typography.labelSmall)
                                }
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

private enum class ReporteTipo(val label: String) {
    RESUMEN("Prestamos por cliente"),
    DETALLADO("Prestamos por cliente - Detallado")
}

private enum class ReporteEstadoFiltro(val label: String) {
    TODOS("Todos"),
    PENDIENTE("Pendiente"),
    PAGADO("Pagado")
}

private data class ReporteClienteFiltro(
    val idCliente: Long?,
    val label: String
)

@Composable
fun ReportesScreen(
    viewModel: AppViewModel,
    isLicenseActive: Boolean
) {
    val context = LocalContext.current
    val clientes by viewModel.clientes.collectAsStateWithLifecycle()
    val prestamos by viewModel.prestamos.collectAsStateWithLifecycle()
    val cuotas by viewModel.cuotas.collectAsStateWithLifecycle()
    val pagos by viewModel.pagos.collectAsStateWithLifecycle()
    val tiposCobro by viewModel.tiposCobro.collectAsStateWithLifecycle()

    var expandedCliente by remember { mutableStateOf(false) }
    var expandedTipoReporte by remember { mutableStateOf(false) }
    var expandedEstadoReporte by remember { mutableStateOf(false) }
    var filtroClienteId by remember { mutableStateOf<Long?>(null) }
    var tipoReporte by remember { mutableStateOf(ReporteTipo.RESUMEN) }
    var estadoReporte by remember { mutableStateOf(ReporteEstadoFiltro.TODOS) }

    val opcionesCliente = remember(clientes) {
        listOf(ReporteClienteFiltro(null, "Todos los clientes")) +
            clientes.map {
                ReporteClienteFiltro(
                    idCliente = it.idCliente,
                    label = "${it.nombre} ${it.apellido} / ${it.documentoIdentidad}"
                )
            }
    }
    val clienteSeleccionado = remember(opcionesCliente, filtroClienteId) {
        opcionesCliente.firstOrNull { it.idCliente == filtroClienteId } ?: opcionesCliente.first()
    }

    val prestamosFiltrados = remember(prestamos, filtroClienteId, estadoReporte) {
        prestamos
            .filter { filtroClienteId == null || it.idCliente == filtroClienteId }
            .filter { prestamo ->
                when (estadoReporte) {
                    ReporteEstadoFiltro.TODOS -> true
                    ReporteEstadoFiltro.PENDIENTE -> prestamo.estadoPrestamo != EstadoPrestamo.PAGADO
                    ReporteEstadoFiltro.PAGADO -> prestamo.estadoPrestamo == EstadoPrestamo.PAGADO
                }
            }
            .sortedByDescending { it.fechaRegistro }
    }
    val cuotasPorPrestamo = remember(cuotas) { cuotas.groupBy { it.idPrestamo } }
    val pagosPorPrestamo = remember(pagos) { pagos.groupBy { it.idPrestamo } }
    val tipoCobroById = remember(tiposCobro) { tiposCobro.associateBy { it.idTipoCobro } }
    val clienteById = remember(clientes) { clientes.associateBy { it.idCliente } }

    val reportePayload = remember(
        tipoReporte,
        prestamosFiltrados,
        cuotasPorPrestamo,
        pagosPorPrestamo,
        tipoCobroById,
        clienteById
    ) {
        when (tipoReporte) {
            ReporteTipo.RESUMEN -> buildReportePrestamosResumenPayload(
                prestamos = prestamosFiltrados,
                cuotasPorPrestamo = cuotasPorPrestamo,
                pagosPorPrestamo = pagosPorPrestamo,
                clienteById = clienteById,
                filtroCliente = clienteSeleccionado.label
            )

            ReporteTipo.DETALLADO -> buildReportePrestamosDetalladoPayload(
                prestamos = prestamosFiltrados,
                cuotasPorPrestamo = cuotasPorPrestamo,
                pagosPorPrestamo = pagosPorPrestamo,
                tipoCobroById = tipoCobroById,
                clienteById = clienteById,
                filtroCliente = clienteSeleccionado.label
            )
        }
    }

    LazyColumn(
        modifier = Modifier
            .fillMaxSize()
            .padding(16.dp),
        verticalArrangement = Arrangement.spacedBy(10.dp)
    ) {
        item {
            Text("Reportes", style = MaterialTheme.typography.headlineSmall)
            DropdownGeneric(
                expanded = expandedCliente,
                onExpandedChange = { expandedCliente = it },
                label = "Cliente",
                selected = clienteSeleccionado.label,
                options = opcionesCliente,
                optionText = { it.label },
                onSelect = { filtroClienteId = it.idCliente }
            )
            Spacer(modifier = Modifier.height(8.dp))
            DropdownGeneric(
                expanded = expandedTipoReporte,
                onExpandedChange = { expandedTipoReporte = it },
                label = "Tipo de reporte",
                selected = tipoReporte.label,
                options = ReporteTipo.entries,
                optionText = { it.label },
                onSelect = { tipoReporte = it }
            )
            Spacer(modifier = Modifier.height(8.dp))
            DropdownGeneric(
                expanded = expandedEstadoReporte,
                onExpandedChange = { expandedEstadoReporte = it },
                label = "Estado",
                selected = estadoReporte.label,
                options = ReporteEstadoFiltro.entries,
                optionText = { it.label },
                onSelect = { estadoReporte = it }
            )
            Spacer(modifier = Modifier.height(8.dp))
            Button(
                onClick = {
                    runCatching {
                        createReportesPdf(context, reportePayload)
                    }.onSuccess { file ->
                        compartirArchivo(context, file, "application/pdf")
                    }.onFailure {
                        Toast.makeText(context, "No se pudo exportar PDF", Toast.LENGTH_SHORT).show()
                    }
                },
                enabled = isLicenseActive && prestamosFiltrados.isNotEmpty()
            ) {
                Text("Generar reporte")
            }
            if (!isLicenseActive) {
                Card(
                    shape = RoundedCornerShape(12.dp),
                    colors = CardDefaults.cardColors(containerColor = Color(0xFFFFE8E8)),
                    modifier = Modifier.fillMaxWidth()
                ) {
                    Text(
                        "Licencia activa requerida para generar reportes",
                        color = Color(0xFF9D2B2B),
                        style = MaterialTheme.typography.bodyMedium,
                        modifier = Modifier.padding(12.dp)
                    )
                }
            } else if (prestamosFiltrados.isEmpty()) {
                Text(
                    "No hay prestamos para el filtro seleccionado.",
                    style = MaterialTheme.typography.bodySmall
                )
            }
        }
    }
}

private fun buildReportePrestamosResumenPayload(
    prestamos: List<com.prestamos.app.data.local.entity.PrestamoEntity>,
    cuotasPorPrestamo: Map<Long, List<com.prestamos.app.data.local.entity.CuotaEntity>>,
    pagosPorPrestamo: Map<Long, List<PagoEntity>>,
    clienteById: Map<Long, ClienteEntity>,
    filtroCliente: String
): ReportPdfPayload {
    val columns = listOf(
        ReportTableColumn("Cliente", 2.6f),
        ReportTableColumn("Documento", 1.6f),
        ReportTableColumn("Telefono", 1.5f),
        ReportTableColumn("Prestamo", 1.0f, TableAlign.CENTER),
        ReportTableColumn("Fecha", 1.5f, TableAlign.CENTER),
        ReportTableColumn("Frecuencia", 1.6f),
        ReportTableColumn("Porcentaje", 1.1f, TableAlign.RIGHT),
        ReportTableColumn("Cuotas", 0.9f, TableAlign.CENTER),
        ReportTableColumn("Capital", 1.4f, TableAlign.RIGHT),
        ReportTableColumn("Capital+Interes", 1.7f, TableAlign.RIGHT),
        ReportTableColumn("Mora", 1.2f, TableAlign.RIGHT),
        ReportTableColumn("Tipo", 1.2f, TableAlign.CENTER),
        ReportTableColumn("Cuotas Pend.", 1.3f, TableAlign.CENTER),
        ReportTableColumn("Saldo", 1.3f, TableAlign.RIGHT)
    )

    val rows = prestamos.map { prestamo ->
        val cliente = clienteById[prestamo.idCliente]
        val cuotas = cuotasPorPrestamo[prestamo.idPrestamo].orEmpty()
        val pagosPrestamo = pagosPorPrestamo[prestamo.idPrestamo].orEmpty()
        val saldoPendiente = cuotas.sumOf { it.saldoPendiente + it.moraPendiente }
        val moraPendiente = cuotas.sumOf { it.moraPendiente }
        val moraCobrada = pagosPrestamo.sumOf { it.moraCobrada }
        val moraTotal = moraPendiente + moraCobrada
        val tipo = if (prestamo.estadoPrestamo == EstadoPrestamo.PAGADO) "Pagado" else "Pendiente"
        val frecuencia = tipoPagoDetalle(prestamo.tipoPago, cuotas)
        val cuotasPendientes = cuotas.count { (it.saldoPendiente + it.moraPendiente) > 0.0 }
        listOf(
            "${cliente?.nombre.orEmpty()} ${cliente?.apellido.orEmpty()}".trim().ifBlank { "-" },
            cliente?.documentoIdentidad ?: "-",
            cliente?.telefono ?: "-",
            "#${prestamo.idPrestamo}",
            prestamo.fechaRegistro.toDateString(),
            frecuencia,
            "${prestamo.interes}%",
                    prestamo.cantidadCuotas.toString(),
                    prestamo.montoPrestado.toMoney(prestamo.moneda),
                    prestamo.montoTotalPrestamo.toMoney(prestamo.moneda),
                    moraTotal.toMoney(prestamo.moneda),
                    tipo,
                    cuotasPendientes.toString(),
                    saldoPendiente.toMoney(prestamo.moneda)
        )
    }

    val generatedAt = SimpleDateFormat("dd/MM/yyyy HH:mm", Locale.getDefault()).format(Date())
    return ReportPdfPayload(
        appName = "AppPrestamos",
        reportType = "Prestamos por cliente",
        filter = filtroCliente,
        generatedAt = generatedAt,
        sections = listOf(
            ReportPdfSection(
                title = "Resumen por prestamo",
                table = ReportTable(columns = columns, rows = rows)
            )
        )
    )
}

private fun buildReportePrestamosDetalladoPayload(
    prestamos: List<com.prestamos.app.data.local.entity.PrestamoEntity>,
    cuotasPorPrestamo: Map<Long, List<com.prestamos.app.data.local.entity.CuotaEntity>>,
    pagosPorPrestamo: Map<Long, List<PagoEntity>>,
    tipoCobroById: Map<Long, TipoCobroEntity>,
    clienteById: Map<Long, ClienteEntity>,
    filtroCliente: String
): ReportPdfPayload {
    val baseColumns = listOf(
        ReportTableColumn("Cliente", 2.6f),
        ReportTableColumn("Documento", 1.6f),
        ReportTableColumn("Telefono", 1.5f),
        ReportTableColumn("Prestamo", 1.0f, TableAlign.CENTER),
        ReportTableColumn("Fecha", 1.5f, TableAlign.CENTER),
        ReportTableColumn("Frecuencia", 1.6f),
        ReportTableColumn("Porcentaje", 1.1f, TableAlign.RIGHT),
        ReportTableColumn("Cuotas", 0.9f, TableAlign.CENTER),
        ReportTableColumn("Capital", 1.4f, TableAlign.RIGHT),
        ReportTableColumn("Capital+Interes", 1.7f, TableAlign.RIGHT),
        ReportTableColumn("Mora", 1.2f, TableAlign.RIGHT),
        ReportTableColumn("Tipo", 1.2f, TableAlign.CENTER),
        ReportTableColumn("Cuotas Pend.", 1.3f, TableAlign.CENTER),
        ReportTableColumn("Saldo", 1.3f, TableAlign.RIGHT)
    )
    val cuotaColumns = listOf(
        ReportTableColumn("Cuota", 0.8f, TableAlign.CENTER),
        ReportTableColumn("Fec.Venc", 1.5f, TableAlign.CENTER),
        ReportTableColumn("Cuota", 1.2f, TableAlign.RIGHT),
        ReportTableColumn("Cuota+Int", 1.4f, TableAlign.RIGHT),
        ReportTableColumn("Mora", 1.1f, TableAlign.RIGHT),
        ReportTableColumn("Cuota+Int+Mora", 1.8f, TableAlign.RIGHT),
        ReportTableColumn("Tipo", 1.0f, TableAlign.CENTER),
        ReportTableColumn("Tipo cobro", 1.4f),
        ReportTableColumn("Fec.Cobro", 1.5f, TableAlign.CENTER)
    )

    val sections = mutableListOf<ReportPdfSection>()
    prestamos.forEach { prestamo ->
        val cliente = clienteById[prestamo.idCliente]
        val cuotas = cuotasPorPrestamo[prestamo.idPrestamo].orEmpty().sortedBy { it.numeroCuota }
        val pagosPrestamo = pagosPorPrestamo[prestamo.idPrestamo].orEmpty()
        val saldoPendiente = cuotas.sumOf { it.saldoPendiente + it.moraPendiente }
        val moraPendiente = cuotas.sumOf { it.moraPendiente }
        val moraCobrada = pagosPrestamo.sumOf { it.moraCobrada }
        val moraTotalPrestamo = moraPendiente + moraCobrada
        val tipo = if (prestamo.estadoPrestamo == EstadoPrestamo.PAGADO) "Pagado" else "Pendiente"
        val frecuencia = tipoPagoDetalle(prestamo.tipoPago, cuotas)
        val cuotaBase = if (prestamo.cantidadCuotas > 0) prestamo.montoPrestado / prestamo.cantidadCuotas else 0.0
        val cuotasPendientes = cuotas.count { (it.saldoPendiente + it.moraPendiente) > 0.0 }

        sections += ReportPdfSection(
            title = "Prestamo #${prestamo.idPrestamo}",
            table = ReportTable(
                columns = baseColumns,
                rows = listOf(
                    listOf(
                        "${cliente?.nombre.orEmpty()} ${cliente?.apellido.orEmpty()}".trim().ifBlank { "-" },
                        cliente?.documentoIdentidad ?: "-",
                        cliente?.telefono ?: "-",
                        "#${prestamo.idPrestamo}",
                        prestamo.fechaRegistro.toDateString(),
                        frecuencia,
                        "${prestamo.interes}%",
                        prestamo.cantidadCuotas.toString(),
                        prestamo.montoPrestado.toMoney(prestamo.moneda),
                        prestamo.montoTotalPrestamo.toMoney(prestamo.moneda),
                        moraTotalPrestamo.toMoney(prestamo.moneda),
                        tipo,
                        cuotasPendientes.toString(),
                        saldoPendiente.toMoney(prestamo.moneda)
                    )
                )
            )
        )

        val cuotaRows = cuotas.map { cuota ->
            val pagosCuota = pagosPrestamo
                .filter { it.idCuota == cuota.idCuota }
                .sortedByDescending { it.fechaPago }
            val moraCobrada = pagosCuota.sumOf { it.moraCobrada }
            val moraTotal = cuota.moraPendiente + moraCobrada
            val cuotaConInteres = cuota.montoCuota
            val cuotaConInteresMora = cuotaConInteres + moraTotal
            val tipoCuota = if ((cuota.saldoPendiente + cuota.moraPendiente) <= 0.0) "Pagado" else "Pendiente"
            val ultimoPago = pagosCuota.firstOrNull()
            val tipoCobro = ultimoPago?.idTipoCobro?.let { tipoCobroById[it]?.nombre } ?: "-"
            val fechaCobro = ultimoPago?.fechaPago?.toDateString() ?: "-"
                listOf(
                    cuota.numeroCuota.toString(),
                    cuota.fechaVencimiento.toDateString(),
                    cuotaBase.toMoney(prestamo.moneda),
                    cuotaConInteres.toMoney(prestamo.moneda),
                    moraTotal.toMoney(prestamo.moneda),
                    cuotaConInteresMora.toMoney(prestamo.moneda),
                    tipoCuota,
                    tipoCobro,
                fechaCobro
            )
        }
        sections += ReportPdfSection(
            title = "Cuotas del prestamo #${prestamo.idPrestamo}",
            table = ReportTable(columns = cuotaColumns, rows = cuotaRows)
        )
    }

    val generatedAt = SimpleDateFormat("dd/MM/yyyy HH:mm", Locale.getDefault()).format(Date())
    return ReportPdfPayload(
        appName = "AppPrestamos",
        reportType = "Prestamos por cliente - Detallado",
        filter = filtroCliente,
        generatedAt = generatedAt,
        sections = sections
    )
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
