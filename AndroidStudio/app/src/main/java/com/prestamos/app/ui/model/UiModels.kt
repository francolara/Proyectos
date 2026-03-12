package com.prestamos.app.ui.model

data class ResumenReporte(
    val totalPrestado: Double = 0.0,
    val totalCobrado: Double = 0.0,
    val totalPendiente: Double = 0.0,
    val prestamosActivos: Int = 0,
    val prestamosPagados: Int = 0,
    val cuotasVencidas: Int = 0
)
