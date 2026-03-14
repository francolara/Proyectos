package com.prestamos.app.util

import com.prestamos.app.data.local.entity.Moneda
import java.text.DecimalFormat
import java.text.DecimalFormatSymbols
import java.time.Instant
import java.time.LocalDate
import java.time.ZoneId
import java.time.format.DateTimeFormatter
import java.util.Locale

private val decimalFormat = DecimalFormat("#,##0.00", DecimalFormatSymbols(Locale.US))
private val dateFormatter = DateTimeFormatter.ofPattern("dd/MM/yyyy")

fun Double.toMoney(moneda: Moneda = Moneda.SOLES): String {
    val simbolo = when (moneda) {
        Moneda.SOLES -> "S/"
        Moneda.DOLARES -> "$"
    }
    return "$simbolo ${decimalFormat.format(this)}"
}

fun Long.toDateString(): String = Instant.ofEpochMilli(this)
    .atZone(ZoneId.systemDefault())
    .toLocalDate()
    .format(dateFormatter)

fun LocalDate.toEpochMillis(): Long = atStartOfDay(ZoneId.systemDefault()).toInstant().toEpochMilli()
