package com.prestamos.app.data.local.entity

import androidx.room.Entity
import androidx.room.Index
import androidx.room.PrimaryKey

@Entity(
    tableName = "clientes",
    indices = [Index(value = ["documentoIdentidad"], unique = true)]
)
data class ClienteEntity(
    @PrimaryKey(autoGenerate = true)
    val idCliente: Long = 0,
    val nombre: String,
    val apellido: String,
    val documentoIdentidad: String,
    val direccion: String,
    val telefono: String,
    val fechaRegistro: Long,
    val fechaModificacion: Long
)
