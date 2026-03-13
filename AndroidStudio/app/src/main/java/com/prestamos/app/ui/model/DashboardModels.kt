package com.prestamos.app.ui.model

import com.prestamos.app.data.local.entity.Moneda

data class DashboardResumen(
    val capitalPrestado: Double = 0.0,
    val saldoPendiente: Double = 0.0,
    val cobradoHoy: Double = 0.0,
    val cuotasVencidas: Int = 0,
    val estadoCuotas: Map<String, Int> = emptyMap(),
    val proximosVencimientos: List<DashboardCuotaItem> = emptyList(),
    val ultimosPagos: List<DashboardPagoItem> = emptyList(),
    val monedaReferencial: Moneda = Moneda.SOLES
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

data class DashboardPagoItem(
    val cliente: String,
    val fechaPago: Long,
    val montoAbono: Double,
    val idPrestamo: Long,
    val idPago: Long,
    val moneda: Moneda
)
