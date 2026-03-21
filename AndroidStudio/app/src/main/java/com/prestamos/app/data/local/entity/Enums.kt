package com.prestamos.app.data.local.entity

enum class TipoPago {
    DIARIO,
    SEMANAL,
    QUINCENAL,
    MENSUAL,
    PERSONALIZADO
}

enum class Moneda(
    val code: String,
    val symbol: String,
    val displayName: String
) {
    SOLES("PEN", "S/", "Sol peruano"),
    DOLARES("USD", "$", "Dolar estadounidense"),
    ARS("ARS", "$", "Peso argentino"),
    BOB("BOB", "Bs", "Boliviano"),
    BRL("BRL", "R$", "Real brasileno"),
    CLP("CLP", "$", "Peso chileno"),
    COP("COP", "$", "Peso colombiano"),
    PYG("PYG", "Gs", "Guarani"),
    UYU("UYU", "\$U", "Peso uruguayo"),
    VES("VES", "Bs", "Bolivar venezolano"),
    GYD("GYD", "G$", "Dolar guyanes"),
    SRD("SRD", "$", "Dolar surinames");

    companion object {
        fun fromCode(code: String?): Moneda? {
            if (code.isNullOrBlank()) return null
            return entries.firstOrNull { it.code.equals(code, ignoreCase = true) }
        }
    }
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
