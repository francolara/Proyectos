package com.prestamos.app.data.local.entity

import androidx.room.Embedded
import androidx.room.Relation

data class ClienteConPrestamos(
    @Embedded val cliente: ClienteEntity,
    @Relation(
        parentColumn = "idCliente",
        entityColumn = "idCliente"
    )
    val prestamos: List<PrestamoEntity>
)

data class PrestamoConCuotas(
    @Embedded val prestamo: PrestamoEntity,
    @Relation(
        parentColumn = "idPrestamo",
        entityColumn = "idPrestamo"
    )
    val cuotas: List<CuotaEntity>
)

data class CuotaConPagos(
    @Embedded val cuota: CuotaEntity,
    @Relation(
        parentColumn = "idCuota",
        entityColumn = "idCuota"
    )
    val pagos: List<PagoEntity>
)
