package com.prestamos.app.data.local.entity

import androidx.room.Entity
import androidx.room.ForeignKey
import androidx.room.Index
import androidx.room.PrimaryKey

@Entity(
    tableName = "prestamos",
    foreignKeys = [
        ForeignKey(
            entity = ClienteEntity::class,
            parentColumns = ["idCliente"],
            childColumns = ["idCliente"],
            onDelete = ForeignKey.RESTRICT
        )
    ],
    indices = [Index(value = ["idCliente"])]
)
data class PrestamoEntity(
    @PrimaryKey(autoGenerate = true)
    val idPrestamo: Long = 0,
    val idCliente: Long,
    val montoPrestado: Double,
    val interes: Double,
    val moneda: Moneda,
    val tipoPago: TipoPago,
    val cantidadCuotas: Int,
    val fechaPrimeraCuota: Long,
    val montoTotalPrestamo: Double,
    val montoCuota: Double,
    val estadoPrestamo: EstadoPrestamo,
    val fechaRegistro: Long,
    val fechaModificacion: Long
)
