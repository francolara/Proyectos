package com.prestamos.app.ui.viewmodel

import android.app.Application
import androidx.lifecycle.AndroidViewModel
import androidx.lifecycle.viewModelScope
import com.prestamos.app.data.local.AppDatabase
import com.prestamos.app.data.local.entity.CuotaEntity
import com.prestamos.app.data.local.entity.EstadoPrestamo
import com.prestamos.app.data.local.entity.Moneda
import com.prestamos.app.data.local.entity.PagoEntity
import com.prestamos.app.data.local.entity.PrestamoEntity
import com.prestamos.app.data.local.entity.TipoPago
import com.prestamos.app.data.local.entity.TipoCobroEntity
import com.prestamos.app.data.repository.PrestamosRepository
import com.prestamos.app.ui.model.ResumenReporte
import kotlinx.coroutines.flow.MutableStateFlow
import kotlinx.coroutines.flow.SharingStarted
import kotlinx.coroutines.flow.StateFlow
import kotlinx.coroutines.flow.combine
import kotlinx.coroutines.flow.flatMapLatest
import kotlinx.coroutines.flow.flowOf
import kotlinx.coroutines.flow.map
import kotlinx.coroutines.flow.stateIn
import kotlinx.coroutines.ExperimentalCoroutinesApi
import kotlinx.coroutines.launch

@OptIn(ExperimentalCoroutinesApi::class)
class AppViewModel(application: Application) : AndroidViewModel(application) {
    private val repository = PrestamosRepository(AppDatabase.getInstance(application))

    val clientes = repository.observarClientes().stateIn(
        scope = viewModelScope,
        started = SharingStarted.WhileSubscribed(5000),
        initialValue = emptyList()
    )

    val prestamos = repository.observarPrestamos().stateIn(
        scope = viewModelScope,
        started = SharingStarted.WhileSubscribed(5000),
        initialValue = emptyList()
    )

    val cuotas = repository.observarCuotas().stateIn(
        scope = viewModelScope,
        started = SharingStarted.WhileSubscribed(5000),
        initialValue = emptyList()
    )

    val pagos = repository.observarPagos().stateIn(
        scope = viewModelScope,
        started = SharingStarted.WhileSubscribed(5000),
        initialValue = emptyList<PagoEntity>()
    )

    val tiposCobro = repository.observarTiposCobro().stateIn(
        scope = viewModelScope,
        started = SharingStarted.WhileSubscribed(5000),
        initialValue = emptyList<TipoCobroEntity>()
    )

    private val clienteSeleccionadoPagos = MutableStateFlow<Long?>(null)
    private val prestamoSeleccionadoPagos = MutableStateFlow<Long?>(null)
    private val prestamoSeleccionadoDetalle = MutableStateFlow<Long?>(null)

    val prestamosClientePagos: StateFlow<List<PrestamoEntity>> = clienteSeleccionadoPagos
        .flatMapLatest { idCliente -> if (idCliente == null) flowOf(emptyList()) else repository.observarPrestamosPorCliente(idCliente) }
        .map { prestamosCliente ->
            prestamosCliente.filter { it.estadoPrestamo == EstadoPrestamo.ACTIVO }
        }
        .stateIn(viewModelScope, SharingStarted.WhileSubscribed(5000), emptyList())

    val cuotasPrestamoPagos: StateFlow<List<CuotaEntity>> = prestamoSeleccionadoPagos
        .flatMapLatest { idPrestamo -> if (idPrestamo == null) flowOf(emptyList()) else repository.observarCuotasPorPrestamo(idPrestamo) }
        .stateIn(viewModelScope, SharingStarted.WhileSubscribed(5000), emptyList())

    val cuotasPrestamoDetalle: StateFlow<List<CuotaEntity>> = prestamoSeleccionadoDetalle
        .flatMapLatest { idPrestamo -> if (idPrestamo == null) flowOf(emptyList()) else repository.observarCuotasPorPrestamo(idPrestamo) }
        .stateIn(viewModelScope, SharingStarted.WhileSubscribed(5000), emptyList())

    val cuotasVencidas = repository.observarCuotasVencidas(System.currentTimeMillis()).stateIn(
        scope = viewModelScope,
        started = SharingStarted.WhileSubscribed(5000),
        initialValue = emptyList()
    )

    val resumenReporte: StateFlow<ResumenReporte> = combine(
        prestamos,
        cuotasVencidas,
        repository.observarTotalCobrado()
    ) { prestamosList, cuotasVencidasList, totalCobrado ->
        val totalPrestado = prestamosList.sumOf { it.montoTotalPrestamo }
        val totalPendiente = prestamosList
            .filter { it.estadoPrestamo != EstadoPrestamo.PAGADO }
            .sumOf { it.montoTotalPrestamo }
        ResumenReporte(
            totalPrestado = totalPrestado,
            totalCobrado = totalCobrado ?: 0.0,
            totalPendiente = totalPendiente,
            prestamosActivos = prestamosList.count { it.estadoPrestamo == EstadoPrestamo.ACTIVO },
            prestamosPagados = prestamosList.count { it.estadoPrestamo == EstadoPrestamo.PAGADO },
            cuotasVencidas = cuotasVencidasList.size
        )
    }.stateIn(
        scope = viewModelScope,
        started = SharingStarted.WhileSubscribed(5000),
        initialValue = ResumenReporte()
    )

    val mensaje = MutableStateFlow<String?>(null)

    fun seleccionarClientePagos(idCliente: Long?) {
        clienteSeleccionadoPagos.value = idCliente
        prestamoSeleccionadoPagos.value = null
    }

    fun seleccionarPrestamoPagos(idPrestamo: Long?) {
        prestamoSeleccionadoPagos.value = idPrestamo
    }

    fun seleccionarPrestamoDetalle(idPrestamo: Long?) {
        prestamoSeleccionadoDetalle.value = idPrestamo
    }

    fun limpiarMensaje() {
        mensaje.value = null
    }

    fun registrarCliente(
        nombre: String,
        apellido: String,
        documento: String,
        direccion: String,
        telefono: String,
        onSuccess: () -> Unit = {}
    ) {
        viewModelScope.launch {
            runCatching {
                val nombreLimpio = cleanSingleLine(nombre)
                val apellidoLimpio = cleanSingleLine(apellido)
                val documentoLimpio = cleanSingleLine(documento)
                val direccionLimpia = cleanSingleLine(direccion)
                val telefonoLimpio = cleanSingleLine(telefono)

                require(nombreLimpio.isNotBlank()) { "Nombres obligatorio" }
                require(apellidoLimpio.isNotBlank()) { "Apellido obligatorio" }
                require(documentoLimpio.isNotBlank()) { "Documento de identidad obligatorio" }

                repository.registrarCliente(
                    nombre = nombreLimpio,
                    apellido = apellidoLimpio,
                    documento = documentoLimpio,
                    direccion = direccionLimpia,
                    telefono = telefonoLimpio
                )
            }.onSuccess {
                onSuccess()
            }.onFailure {
                mensaje.value = it.message ?: "No se pudo registrar cliente"
            }
        }
    }

    fun registrarPrestamo(
        idCliente: Long,
        monto: String,
        interes: String,
        moneda: Moneda,
        tipoPago: TipoPago,
        intervaloDiasPersonalizado: String,
        cuotas: String,
        fechaPrimeraCuota: Long,
        onSuccess: () -> Unit = {}
    ) {
        viewModelScope.launch {
            runCatching {
                val montoDouble = monto.toDoubleOrNull() ?: error("Monto invalido")
                val interesDouble = interes.toDoubleOrNull() ?: error("Interes invalido")
                val cuotasInt = cuotas.toIntOrNull() ?: error("Cuotas invalidas")
                val intervaloPersonalizadoInt = if (tipoPago == TipoPago.PERSONALIZADO) {
                    intervaloDiasPersonalizado.toIntOrNull() ?: error("Intervalo personalizado invalido")
                } else {
                    null
                }
                require(montoDouble > 0.0) { "Monto debe ser mayor a 0" }
                require(interesDouble >= 0.0) { "Interes debe ser mayor o igual a 0" }
                require(cuotasInt > 0) { "Cuotas debe ser mayor a 0" }
                if (tipoPago == TipoPago.PERSONALIZADO) {
                    require((intervaloPersonalizadoInt ?: 0) > 0) { "Intervalo personalizado invalido" }
                }
                repository.registrarPrestamo(
                    idCliente = idCliente,
                    monto = montoDouble,
                    interesPorcentaje = interesDouble,
                    moneda = moneda,
                    tipoPago = tipoPago,
                    intervaloDiasPersonalizado = intervaloPersonalizadoInt,
                    cantidadCuotas = cuotasInt,
                    fechaPrimeraCuota = fechaPrimeraCuota
                )
            }.onSuccess {
                onSuccess()
            }.onFailure {
                mensaje.value = it.message ?: "No se pudo registrar prestamo"
            }
        }
    }

    fun registrarPago(
        idPrestamo: Long,
        idCuota: Long,
        idTipoCobro: Long?,
        montoAbono: String,
        onSuccess: () -> Unit = {}
    ) {
        viewModelScope.launch {
            runCatching {
                val tipoCobroId = idTipoCobro ?: error("Selecciona un tipo de cobro")
                repository.registrarPago(
                    idPrestamo = idPrestamo,
                    idCuota = idCuota,
                    idTipoCobro = tipoCobroId,
                    montoAbono = montoAbono.toDoubleOrNull() ?: error("Monto abonado invalido"),
                    observacion = null
                )
            }.onSuccess {
                onSuccess()
            }.onFailure {
                mensaje.value = it.message ?: "No se pudo registrar pago"
            }
        }
    }

    fun registrarTipoCobro(nombre: String, onSuccess: () -> Unit = {}) {
        viewModelScope.launch {
            runCatching {
                repository.registrarTipoCobro(nombre)
            }.onSuccess {
                mensaje.value = "Tipo de cobro registrado"
                onSuccess()
            }.onFailure {
                mensaje.value = it.message ?: "No se pudo registrar tipo de cobro"
            }
        }
    }

    fun eliminarTipoCobro(idTipoCobro: Long, onSuccess: () -> Unit = {}) {
        viewModelScope.launch {
            runCatching {
                repository.eliminarTipoCobroSiNoTienePagos(idTipoCobro)
            }.onSuccess {
                mensaje.value = "Tipo de cobro eliminado"
                onSuccess()
            }.onFailure {
                mensaje.value = it.message ?: "No se pudo eliminar tipo de cobro"
            }
        }
    }

    fun eliminarPrestamo(idPrestamo: Long) {
        viewModelScope.launch {
            runCatching {
                repository.eliminarPrestamoSiNoTienePagos(idPrestamo)
            }.onSuccess {
                mensaje.value = "Prestamo eliminado correctamente"
            }.onFailure {
                mensaje.value = it.message ?: "No se pudo eliminar el prestamo"
            }
        }
    }

    fun eliminarPago(idPago: Long, onSuccess: () -> Unit = {}) {
        viewModelScope.launch {
            runCatching {
                repository.eliminarPagoSiEsUltimo(idPago)
            }.onSuccess {
                mensaje.value = "Pago eliminado correctamente"
                onSuccess()
            }.onFailure {
                mensaje.value = it.message ?: "No se pudo eliminar el pago"
            }
        }
    }

    fun aplicarMoraCuotaVencida(
        idCuota: Long,
        montoMora: String,
        onSuccess: () -> Unit = {}
    ) {
        viewModelScope.launch {
            runCatching {
                val monto = montoMora.toDoubleOrNull() ?: error("Monto de mora invalido")
                repository.aplicarMoraManualCuotaVencida(
                    idCuota = idCuota,
                    montoMora = monto
                )
            }.onSuccess {
                mensaje.value = "Mora aplicada correctamente"
                onSuccess()
            }.onFailure {
                mensaje.value = it.message ?: "No se pudo aplicar mora"
            }
        }
    }

    fun actualizarCliente(
        idCliente: Long,
        nombre: String,
        apellido: String,
        direccion: String,
        telefono: String
    ) {
        viewModelScope.launch {
            runCatching {
                val nombreLimpio = cleanSingleLine(nombre)
                val apellidoLimpio = cleanSingleLine(apellido)
                val direccionLimpia = cleanSingleLine(direccion)
                val telefonoLimpio = cleanSingleLine(telefono)
                require(nombreLimpio.isNotBlank()) { "Nombres obligatorio" }
                require(apellidoLimpio.isNotBlank()) { "Apellido obligatorio" }
                repository.actualizarCliente(
                    idCliente = idCliente,
                    nombre = nombreLimpio,
                    apellido = apellidoLimpio,
                    direccion = direccionLimpia,
                    telefono = telefonoLimpio
                )
            }.onSuccess {
                mensaje.value = "Cliente actualizado correctamente"
            }.onFailure {
                mensaje.value = it.message ?: "No se pudo actualizar cliente"
            }
        }
    }

    fun eliminarCliente(idCliente: Long) {
        viewModelScope.launch {
            runCatching {
                repository.eliminarClienteSiNoTienePrestamos(idCliente)
            }.onSuccess {
                mensaje.value = "Cliente eliminado correctamente"
            }.onFailure {
                mensaje.value = it.message ?: "No se pudo eliminar cliente"
            }
        }
    }

    private fun cleanSingleLine(value: String): String =
        value.replace("\n", " ").replace("\r", " ").trim()
}
