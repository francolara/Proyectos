package com.prestamos.app.data.local.dao

import androidx.room.Dao
import androidx.room.Insert
import androidx.room.OnConflictStrategy
import androidx.room.Query
import androidx.room.Update
import com.prestamos.app.data.local.entity.CuotaEntity
import kotlinx.coroutines.flow.Flow

@Dao
interface CuotaDao {
    @Insert(onConflict = OnConflictStrategy.ABORT)
    suspend fun insertarCuotas(cuotas: List<CuotaEntity>)

    @Update
    suspend fun actualizar(cuota: CuotaEntity)

    @Query("SELECT * FROM cuotas WHERE idPrestamo = :idPrestamo ORDER BY numeroCuota")
    fun listarPorPrestamo(idPrestamo: Long): Flow<List<CuotaEntity>>
}
