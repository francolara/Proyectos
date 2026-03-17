package com.prestamos.app.ui.viewmodel

import android.app.Application
import androidx.lifecycle.AndroidViewModel
import androidx.lifecycle.viewModelScope
import com.prestamos.app.data.local.AppDatabase
import com.prestamos.app.data.local.entity.EstadoCuota
import com.prestamos.app.data.local.entity.EstadoPrestamo
import com.prestamos.app.data.local.entity.Moneda
import com.prestamos.app.data.repository.PrestamosRepository
import com.prestamos.app.ui.model.DashboardCuotaDetalleItem
import com.prestamos.app.ui.model.DashboardCuotaItem
import com.prestamos.app.ui.model.DashboardGananciaPrestamoItem
import com.prestamos.app.ui.model.DashboardPagoItem
import com.prestamos.app.ui.model.DashboardPrestamoDetalleItem
import com.prestamos.app.ui.model.DashboardResumen
import kotlinx.coroutines.flow.SharingStarted
import kotlinx.coroutines.flow.StateFlow
import kotlinx.coroutines.flow.combine
import kotlinx.coroutines.flow.stateIn
import java.time.LocalDate
import java.time.ZoneId

class DashboardViewModel(application: Application) : AndroidViewModel(application) {
    private val repository = PrestamosRepository(AppDatabase.getInstance(application))

    val uiState: StateFlow<DashboardResumen> = combine(
        repository.observarClientes(),
        repository.observarPrestamos(),
        repository.observarCuotas(),
        repository.observarPagos()
    ) { clientes, prestamos, cuotas, pagos ->
        val prestamoById = prestamos.associateBy { it.idPrestamo }
        val clienteById = clientes.associateBy { it.idCliente }
        val cuotasByPrestamo = cuotas.groupBy { it.idPrestamo }
        val pagosByPrestamo = pagos.groupBy { it.idPrestamo }
        val cuotaById = cuotas.associateBy { it.idCuota }

        val now = System.currentTimeMillis()
        val startToday = LocalDate.now().atStartOfDay(ZoneId.systemDefault()).toInstant().toEpochMilli()
        val endToday = LocalDate.now().plusDays(1).atStartOfDay(ZoneId.systemDefault()).toInstant().toEpochMilli() - 1

        val capitalPrestado = prestamos.sumOf { it.montoPrestado }
        val saldoPendiente = cuotas.sumOf { it.saldoPendiente }
        val cobradoHoy = pagos.filter { it.fechaPago in startToday..endToday }.sumOf { it.montoAbono }
        val cuotasVencidas = cuotas.count { it.fechaVencimiento <= endToday && it.saldoPendiente > 0.0 }

        val estadoCuotas = mapOf(
            "Pagadas" to cuotas.count { it.estadoCuota == EstadoCuota.PAGADO },
            "Pendientes" to cuotas.count { it.estadoCuota == EstadoCuota.PENDIENTE },
            "Parciales" to cuotas.count { it.estadoCuota == EstadoCuota.PARCIAL },
            "Vencidas" to cuotas.count { it.estadoCuota == EstadoCuota.VENCIDO || (it.fechaVencimiento < now && it.saldoPendiente > 0.0) }
        )

        val prestamosActivosDetalle = prestamos
            .mapNotNull { prestamo ->
                val cliente = clienteById[prestamo.idCliente]
                val cuotasPrestamo = cuotasByPrestamo[prestamo.idPrestamo].orEmpty()
                val saldoPrestamo = cuotasPrestamo.sumOf { it.saldoPendiente }
                val cuotasPendientes = cuotasPrestamo.count { it.saldoPendiente > 0.0 }
                if (prestamo.estadoPrestamo != EstadoPrestamo.ACTIVO && cuotasPendientes == 0) return@mapNotNull null
                DashboardPrestamoDetalleItem(
                    cliente = "${cliente?.nombre.orEmpty()} ${cliente?.apellido.orEmpty()}".trim().ifBlank { "-" },
                    idPrestamo = prestamo.idPrestamo,
                    montoPrestado = prestamo.montoPrestado,
                    montoTotalConInteres = prestamo.montoTotalPrestamo,
                    montoCobrado = pagosByPrestamo[prestamo.idPrestamo].orEmpty().sumOf { it.montoAbono },
                    saldoPendiente = saldoPrestamo,
                    totalCuotas = prestamo.cantidadCuotas,
                    cuotasPendientes = cuotasPendientes,
                    moneda = prestamo.moneda
                )
            }
            .sortedByDescending { it.saldoPendiente }

        val prestamosCapitalDetalle = prestamos
            .filter { it.estadoPrestamo == EstadoPrestamo.ACTIVO || it.estadoPrestamo == EstadoPrestamo.PAGADO }
            .map { prestamo ->
                val cliente = clienteById[prestamo.idCliente]
                val cuotasPrestamo = cuotasByPrestamo[prestamo.idPrestamo].orEmpty()
                DashboardPrestamoDetalleItem(
                    cliente = "${cliente?.nombre.orEmpty()} ${cliente?.apellido.orEmpty()}".trim().ifBlank { "-" },
                    idPrestamo = prestamo.idPrestamo,
                    montoPrestado = prestamo.montoPrestado,
                    montoTotalConInteres = prestamo.montoTotalPrestamo,
                    montoCobrado = pagosByPrestamo[prestamo.idPrestamo].orEmpty().sumOf { it.montoAbono },
                    saldoPendiente = cuotasPrestamo.sumOf { it.saldoPendiente },
                    totalCuotas = prestamo.cantidadCuotas,
                    cuotasPendientes = cuotasPrestamo.count { it.saldoPendiente > 0.0 },
                    moneda = prestamo.moneda
                )
            }
            .sortedWith(compareByDescending<DashboardPrestamoDetalleItem> { it.montoPrestado }.thenByDescending { it.idPrestamo })

        val gananciasPrestamosPagados = prestamos
            .filter { it.estadoPrestamo == EstadoPrestamo.PAGADO }
            .map { prestamo ->
                val cliente = clienteById[prestamo.idCliente]
                val montoCobrado = prestamo.montoTotalPrestamo
                DashboardGananciaPrestamoItem(
                    cliente = "${cliente?.nombre.orEmpty()} ${cliente?.apellido.orEmpty()}".trim().ifBlank { "-" },
                    idPrestamo = prestamo.idPrestamo,
                    montoPrestado = prestamo.montoPrestado,
                    montoCobrado = montoCobrado,
                    ganancia = montoCobrado - prestamo.montoPrestado,
                    moneda = prestamo.moneda
                )
            }
            .sortedByDescending { it.ganancia }

        val gananciaAcumulada = gananciasPrestamosPagados.sumOf { it.ganancia }

        val cuotasPendientesDetalle = cuotas
            .filter { it.saldoPendiente > 0.0 }
            .sortedBy { it.fechaVencimiento }
            .map { cuota ->
                val prestamo = prestamoById[cuota.idPrestamo]
                val cliente = clienteById[prestamo?.idCliente]
                DashboardCuotaDetalleItem(
                    cliente = "${cliente?.nombre.orEmpty()} ${cliente?.apellido.orEmpty()}".trim().ifBlank { "-" },
                    idPrestamo = cuota.idPrestamo,
                    numeroCuota = cuota.numeroCuota,
                    fechaVencimiento = cuota.fechaVencimiento,
                    saldoPendiente = cuota.saldoPendiente,
                    estado = cuota.estadoCuota,
                    moneda = prestamo?.moneda ?: Moneda.SOLES
                )
            }

        val cuotasVencidasDetalle = cuotas
            .filter { it.fechaVencimiento <= endToday && it.saldoPendiente > 0.0 }
            .sortedBy { it.fechaVencimiento }
            .map { cuota ->
                val prestamo = prestamoById[cuota.idPrestamo]
                val cliente = clienteById[prestamo?.idCliente]
                DashboardCuotaDetalleItem(
                    cliente = "${cliente?.nombre.orEmpty()} ${cliente?.apellido.orEmpty()}".trim().ifBlank { "-" },
                    idPrestamo = cuota.idPrestamo,
                    numeroCuota = cuota.numeroCuota,
                    fechaVencimiento = cuota.fechaVencimiento,
                    saldoPendiente = cuota.saldoPendiente,
                    estado = cuota.estadoCuota,
                    moneda = prestamo?.moneda ?: Moneda.SOLES
                )
            }

        val pagosHoyDetalle = pagos
            .filter { it.fechaPago in startToday..endToday }
            .sortedByDescending { it.fechaPago }
            .map { pago ->
                val prestamo = prestamoById[pago.idPrestamo]
                val cliente = clienteById[prestamo?.idCliente]
                DashboardPagoItem(
                    cliente = "${cliente?.nombre.orEmpty()} ${cliente?.apellido.orEmpty()}".trim().ifBlank { "-" },
                    fechaPago = pago.fechaPago,
                    montoAbono = pago.montoAbono,
                    idPrestamo = pago.idPrestamo,
                    numeroCuota = cuotaById[pago.idCuota]?.numeroCuota ?: 0,
                    idPago = pago.idPago,
                    moneda = prestamo?.moneda ?: Moneda.SOLES
                )
            }

        val proximosVencimientos = cuotasPendientesDetalle
            .filter { it.fechaVencimiento >= startToday }
            .take(5)
            .map {
                DashboardCuotaItem(
                    cliente = it.cliente,
                    numeroCuota = it.numeroCuota,
                    fechaVencimiento = it.fechaVencimiento,
                    saldoPendiente = it.saldoPendiente,
                    estado = it.estado.name,
                    idPrestamo = it.idPrestamo,
                    idCuota = 0,
                    moneda = it.moneda
                )
            }

        val ultimosPagos = pagos
            .sortedByDescending { it.fechaPago }
            .take(5)
            .map { pago ->
                val prestamo = prestamoById[pago.idPrestamo]
                val cliente = clienteById[prestamo?.idCliente]
                DashboardPagoItem(
                    cliente = "${cliente?.nombre.orEmpty()} ${cliente?.apellido.orEmpty()}".trim().ifBlank { "-" },
                    fechaPago = pago.fechaPago,
                    montoAbono = pago.montoAbono,
                    idPrestamo = pago.idPrestamo,
                    numeroCuota = cuotaById[pago.idCuota]?.numeroCuota ?: 0,
                    idPago = pago.idPago,
                    moneda = prestamo?.moneda ?: Moneda.SOLES
                )
            }

        DashboardResumen(
            capitalPrestado = capitalPrestado,
            saldoPendiente = saldoPendiente,
            cobradoHoy = cobradoHoy,
            cuotasVencidas = cuotasVencidas,
            estadoCuotas = estadoCuotas,
            proximosVencimientos = proximosVencimientos,
            ultimosPagos = ultimosPagos,
            prestamosCapitalDetalle = prestamosCapitalDetalle,
            prestamosActivosDetalle = prestamosActivosDetalle,
            cuotasPendientesDetalle = cuotasPendientesDetalle,
            cuotasVencidasDetalle = cuotasVencidasDetalle,
            pagosHoyDetalle = pagosHoyDetalle,
            gananciasPrestamosPagados = gananciasPrestamosPagados,
            gananciaAcumulada = gananciaAcumulada,
            monedaReferencial = prestamos.firstOrNull()?.moneda ?: Moneda.SOLES
        )
    }.stateIn(
        scope = viewModelScope,
        started = SharingStarted.WhileSubscribed(5000),
        initialValue = DashboardResumen()
    )
}
