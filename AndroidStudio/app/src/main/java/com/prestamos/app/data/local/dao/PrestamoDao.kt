package com.prestamos.app.data.local.dao

import androidx.room.Dao
import androidx.room.Insert
import androidx.room.OnConflictStrategy
import androidx.room.Query
import com.prestamos.app.data.local.entity.PrestamoEntity
import kotlinx.coroutines.flow.Flow

@Dao
interface PrestamoDao {
    @Insert(onConflict = OnConflictStrategy.ABORT)
    suspend fun insertar(prestamo: PrestamoEntity): Long

    @Query("SELECT * FROM prestamos WHERE idCliente = :idCliente ORDER BY fechaRegistro DESC")
    fun listarPorCliente(idCliente: Long): Flow<List<PrestamoEntity>>

    @Query("SELECT * FROM prestamos ORDER BY fechaRegistro DESC")
    fun listarTodos(): Flow<List<PrestamoEntity>>
}
