package com.prestamos.app.data.local.entity

enum class TipoPago {
    DIARIO,
    SEMANAL,
    MENSUAL
}

enum class Moneda {
    SOLES,
    DOLARES
}

enum class EstadoPrestamo {
    ACTIVO,
    PAGADO,
    VENCIDO
}

enum class EstadoCuota {
    PENDIENTE,
    PARCIAL,
    PAGADO,
    VENCIDO
}
