package com.prestamos.app.data.local

import androidx.room.TypeConverter
import com.prestamos.app.data.local.entity.EstadoCuota
import com.prestamos.app.data.local.entity.EstadoPrestamo
import com.prestamos.app.data.local.entity.TipoPago

class Converters {
    @TypeConverter
    fun fromTipoPago(value: TipoPago): String = value.name

    @TypeConverter
    fun toTipoPago(value: String): TipoPago = TipoPago.valueOf(value)

    @TypeConverter
    fun fromEstadoPrestamo(value: EstadoPrestamo): String = value.name

    @TypeConverter
    fun toEstadoPrestamo(value: String): EstadoPrestamo = EstadoPrestamo.valueOf(value)

    @TypeConverter
    fun fromEstadoCuota(value: EstadoCuota): String = value.name

    @TypeConverter
    fun toEstadoCuota(value: String): EstadoCuota = EstadoCuota.valueOf(value)
}
