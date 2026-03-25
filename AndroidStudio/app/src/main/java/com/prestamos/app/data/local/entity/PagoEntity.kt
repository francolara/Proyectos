package com.prestamos.app.data.local.entity

import androidx.room.Entity
import androidx.room.ForeignKey
import androidx.room.Index
import androidx.room.PrimaryKey

@Entity(
    tableName = "pagos",
    foreignKeys = [
        ForeignKey(
            entity = PrestamoEntity::class,
            parentColumns = ["idPrestamo"],
            childColumns = ["idPrestamo"],
            onDelete = ForeignKey.CASCADE
        ),
        ForeignKey(
            entity = CuotaEntity::class,
            parentColumns = ["idCuota"],
            childColumns = ["idCuota"],
            onDelete = ForeignKey.CASCADE
        )
    ],
    indices = [
        Index(value = ["idPrestamo"]),
        Index(value = ["idCuota"]),
        Index(value = ["idTipoCobro"])
    ]
)
data class PagoEntity(
    @PrimaryKey(autoGenerate = true)
    val idPago: Long = 0,
    val idPrestamo: Long,
    val idCuota: Long,
    val idTipoCobro: Long?,
    val fechaPago: Long,
    val montoAbono: Double,
    val moraCobrada: Double = 0.0,
    val observacion: String?,
    val fechaRegistro: Long,
    val fechaModificacion: Long
)
