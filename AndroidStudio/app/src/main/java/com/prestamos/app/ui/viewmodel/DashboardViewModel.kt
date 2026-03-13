package com.prestamos.app.ui.viewmodel

import android.app.Application
import androidx.lifecycle.AndroidViewModel
import androidx.lifecycle.viewModelScope
import com.prestamos.app.data.local.AppDatabase
import com.prestamos.app.data.local.entity.EstadoCuota
import com.prestamos.app.data.repository.PrestamosRepository
import com.prestamos.app.ui.model.DashboardCuotaItem
import com.prestamos.app.ui.model.DashboardPagoItem
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

        val now = System.currentTimeMillis()
        val startToday = LocalDate.now().atStartOfDay(ZoneId.systemDefault()).toInstant().toEpochMilli()
        val endToday = LocalDate.now().plusDays(1).atStartOfDay(ZoneId.systemDefault()).toInstant().toEpochMilli() - 1

        val capitalPrestado = prestamos.sumOf { it.montoPrestado }
        val saldoPendiente = cuotas.sumOf { it.saldoPendiente }
        val cobradoHoy = pagos.filter { it.fechaPago in startToday..endToday }.sumOf { it.montoAbono }
        val cuotasVencidas = cuotas.count { it.fechaVencimiento < now && it.saldoPendiente > 0.0 }

        val estadoCuotas = mapOf(
            "Pagadas" to cuotas.count { it.estadoCuota == EstadoCuota.PAGADO },
            "Pendientes" to cuotas.count { it.estadoCuota == EstadoCuota.PENDIENTE },
            "Parciales" to cuotas.count { it.estadoCuota == EstadoCuota.PARCIAL },
            "Vencidas" to cuotas.count { it.estadoCuota == EstadoCuota.VENCIDO || (it.fechaVencimiento < now && it.saldoPendiente > 0.0) }
        )

        val proximosVencimientos = cuotas
            .filter { it.saldoPendiente > 0.0 && it.fechaVencimiento >= startToday }
            .sortedBy { it.fechaVencimiento }
            .take(5)
            .map { cuota ->
                val prestamo = prestamoById[cuota.idPrestamo]
                val cliente = clienteById[prestamo?.idCliente]
                DashboardCuotaItem(
                    cliente = "${cliente?.nombre.orEmpty()} ${cliente?.apellido.orEmpty()}".trim().ifBlank { "-" },
                    numeroCuota = cuota.numeroCuota,
                    fechaVencimiento = cuota.fechaVencimiento,
                    saldoPendiente = cuota.saldoPendiente,
                    estado = cuota.estadoCuota.name,
                    idPrestamo = cuota.idPrestamo,
                    idCuota = cuota.idCuota,
                    moneda = prestamo?.moneda ?: com.prestamos.app.data.local.entity.Moneda.SOLES
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
                    idPago = pago.idPago,
                    moneda = prestamo?.moneda ?: com.prestamos.app.data.local.entity.Moneda.SOLES
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
            monedaReferencial = prestamos.firstOrNull()?.moneda ?: com.prestamos.app.data.local.entity.Moneda.SOLES
        )
    }.stateIn(
        scope = viewModelScope,
        started = SharingStarted.WhileSubscribed(5000),
        initialValue = DashboardResumen()
    )
}
