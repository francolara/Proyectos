package com.prestamos.app.ui.model

import com.prestamos.app.data.local.entity.EstadoCuota
import com.prestamos.app.data.local.entity.Moneda

data class DashboardResumen(
    val capitalPrestado: Double = 0.0,
    val saldoPendiente: Double = 0.0,
    val cobradoHoy: Double = 0.0,
    val cuotasVencidas: Int = 0,
    val estadoCuotas: Map<String, Int> = emptyMap(),
    val proximosVencimientos: List<DashboardCuotaItem> = emptyList(),
    val ultimosPagos: List<DashboardPagoItem> = emptyList(),
    val prestamosCapitalDetalle: List<DashboardPrestamoDetalleItem> = emptyList(),
    val prestamosActivosDetalle: List<DashboardPrestamoDetalleItem> = emptyList(),
    val cuotasPendientesDetalle: List<DashboardCuotaDetalleItem> = emptyList(),
    val cuotasVencidasDetalle: List<DashboardCuotaDetalleItem> = emptyList(),
    val pagosHoyDetalle: List<DashboardPagoItem> = emptyList(),
    val gananciasPrestamosPagados: List<DashboardGananciaPrestamoItem> = emptyList(),
    val gananciaAcumulada: Double = 0.0,
    val monedaReferencial: Moneda = Moneda.SOLES
)

data class DashboardGananciaPrestamoItem(
    val cliente: String,
    val idPrestamo: Long,
    val montoPrestado: Double,
    val montoCobrado: Double,
    val ganancia: Double,
    val moneda: Moneda
)

data class DashboardPrestamoDetalleItem(
    val cliente: String,
    val idPrestamo: Long,
    val montoPrestado: Double,
    val montoTotalConInteres: Double,
    val montoCobrado: Double,
    val saldoPendiente: Double,
    val totalCuotas: Int,
    val cuotasPendientes: Int,
    val moneda: Moneda
)

data class DashboardCuotaItem(
    val cliente: String,
    val numeroCuota: Int,
    val fechaVencimiento: Long,
    val saldoPendiente: Double,
    val estado: String,
    val idPrestamo: Long,
    val idCuota: Long,
    val moneda: Moneda
)

data class DashboardCuotaDetalleItem(
    val cliente: String,
    val idPrestamo: Long,
    val numeroCuota: Int,
    val fechaVencimiento: Long,
    val saldoPendiente: Double,
    val estado: EstadoCuota,
    val moneda: Moneda
)

data class DashboardPagoItem(
    val cliente: String,
    val fechaPago: Long,
    val montoAbono: Double,
    val idPrestamo: Long,
    val numeroCuota: Int,
    val idPago: Long,
    val moneda: Moneda
)
