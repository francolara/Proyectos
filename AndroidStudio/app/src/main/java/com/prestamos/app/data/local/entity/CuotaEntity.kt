package com.prestamos.app.data.local.entity

import androidx.room.Entity
import androidx.room.ForeignKey
import androidx.room.Index
import androidx.room.PrimaryKey

@Entity(
    tableName = "cuotas",
    foreignKeys = [
        ForeignKey(
            entity = PrestamoEntity::class,
            parentColumns = ["idPrestamo"],
            childColumns = ["idPrestamo"],
            onDelete = ForeignKey.CASCADE
        )
    ],
    indices = [Index(value = ["idPrestamo"])]
)
data class CuotaEntity(
    @PrimaryKey(autoGenerate = true)
    val idCuota: Long = 0,
    val idPrestamo: Long,
    val numeroCuota: Int,
    val fechaVencimiento: Long,
    val montoCuota: Double,
    val montoPagado: Double,
    val saldoPendiente: Double,
    val estadoCuota: EstadoCuota,
    val fechaRegistro: Long,
    val fechaModificacion: Long
)
