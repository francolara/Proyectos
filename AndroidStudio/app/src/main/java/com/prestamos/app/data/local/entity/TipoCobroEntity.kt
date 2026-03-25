package com.prestamos.app.data.local.entity

import androidx.room.Entity
import androidx.room.Index
import androidx.room.PrimaryKey

@Entity(
    tableName = "tipos_cobro",
    indices = [Index(value = ["nombre"], unique = true)]
)
data class TipoCobroEntity(
    @PrimaryKey(autoGenerate = true)
    val idTipoCobro: Long = 0,
    val nombre: String,
    val fechaRegistro: Long,
    val fechaModificacion: Long
)

