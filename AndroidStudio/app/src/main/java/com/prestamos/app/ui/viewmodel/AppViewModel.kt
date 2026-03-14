package com.prestamos.app.ui.viewmodel

import android.app.Application
import androidx.lifecycle.AndroidViewModel
import androidx.lifecycle.viewModelScope
import com.prestamos.app.data.local.AppDatabase
import com.prestamos.app.data.local.entity.CuotaEntity
import com.prestamos.app.data.local.entity.EstadoPrestamo
import com.prestamos.app.data.local.entity.Moneda
import com.prestamos.app.data.local.entity.PrestamoEntity
import com.prestamos.app.data.local.entity.TipoPago
import com.prestamos.app.data.repository.PrestamosRepository
import com.prestamos.app.ui.model.ResumenReporte
import kotlinx.coroutines.flow.MutableStateFlow
import kotlinx.coroutines.flow.SharingStarted
import kotlinx.coroutines.flow.StateFlow
import kotlinx.coroutines.flow.combine
import kotlinx.coroutines.flow.flatMapLatest
import kotlinx.coroutines.flow.flowOf
import kotlinx.coroutines.flow.stateIn
import kotlinx.coroutines.launch

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

    private val clienteSeleccionadoPagos = MutableStateFlow<Long?>(null)
    private val prestamoSeleccionadoPagos = MutableStateFlow<Long?>(null)
    private val prestamoSeleccionadoDetalle = MutableStateFlow<Long?>(null)

    val prestamosClientePagos: StateFlow<List<PrestamoEntity>> = clienteSeleccionadoPagos
        .flatMapLatest { idCliente -> if (idCliente == null) flowOf(emptyList()) else repository.observarPrestamosPorCliente(idCliente) }
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

    fun registrarCliente(nombre: String, apellido: String, documento: String, direccion: String, telefono: String) {
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
                mensaje.value = "Cliente registrado"
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
        cuotas: String,
        fechaPrimeraCuota: Long
    ) {
        viewModelScope.launch {
            runCatching {
                val montoDouble = monto.toDoubleOrNull() ?: error("Monto inválido")
                val interesDouble = interes.toDoubleOrNull() ?: error("Interés inválido")
                val cuotasInt = cuotas.toIntOrNull() ?: error("Cuotas inválidas")
                require(montoDouble > 0.0) { "Monto debe ser mayor a 0" }
                require(interesDouble >= 0.0) { "Interés debe ser mayor o igual a 0" }
                require(cuotasInt > 0) { "Cuotas debe ser mayor a 0" }
                repository.registrarPrestamo(
                    idCliente = idCliente,
                    monto = montoDouble,
                    interesPorcentaje = interesDouble,
                    moneda = moneda,
                    tipoPago = tipoPago,
                    cantidadCuotas = cuotasInt,
                    fechaPrimeraCuota = fechaPrimeraCuota
                )
            }.onSuccess {
                mensaje.value = "Préstamo registrado con cuotas generadas"
            }.onFailure {
                mensaje.value = it.message ?: "No se pudo registrar préstamo"
            }
        }
    }

    fun registrarPago(idPrestamo: Long, idCuota: Long, montoAbono: String) {
        viewModelScope.launch {
            runCatching {
                repository.registrarPago(
                    idPrestamo = idPrestamo,
                    idCuota = idCuota,
                    montoAbono = montoAbono.toDoubleOrNull() ?: error("Monto abonado inválido"),
                    observacion = null
                )
            }.onSuccess {
                mensaje.value = "Pago registrado"
            }.onFailure {
                mensaje.value = it.message ?: "No se pudo registrar pago"
            }
        }
    }

    private fun cleanSingleLine(value: String): String =
        value.replace("\n", " ").replace("\r", " ").trim()
}
