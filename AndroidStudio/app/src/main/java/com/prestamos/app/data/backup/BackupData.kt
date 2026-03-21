package com.prestamos.app.data.backup

import com.prestamos.app.data.local.entity.ClienteEntity
import com.prestamos.app.data.local.entity.CuotaEntity
import com.prestamos.app.data.local.entity.PagoEntity
import com.prestamos.app.data.local.entity.PrestamoEntity

data class BackupData(
    val version: Int,
    val fechaBackup: Long,
    val clientes: List<ClienteEntity>,
    val prestamos: List<PrestamoEntity>,
    val cuotas: List<CuotaEntity>,
    val pagos: List<PagoEntity>
)
