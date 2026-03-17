package com.prestamos.app.data.local.dao

import androidx.room.Dao
import androidx.room.Insert
import androidx.room.OnConflictStrategy
import androidx.room.Query
import androidx.room.Update
import com.prestamos.app.data.local.entity.PrestamoEntity
import kotlinx.coroutines.flow.Flow

@Dao
interface PrestamoDao {
    @Insert(onConflict = OnConflictStrategy.ABORT)
    suspend fun insertar(prestamo: PrestamoEntity): Long

    @Update
    suspend fun actualizar(prestamo: PrestamoEntity)

    @Query("SELECT * FROM prestamos WHERE idCliente = :idCliente ORDER BY fechaRegistro DESC")
    fun listarPorCliente(idCliente: Long): Flow<List<PrestamoEntity>>

    @Query("SELECT * FROM prestamos ORDER BY fechaRegistro DESC")
    fun listarTodos(): Flow<List<PrestamoEntity>>

    @Query("SELECT * FROM prestamos WHERE idPrestamo = :idPrestamo LIMIT 1")
    suspend fun obtenerPorId(idPrestamo: Long): PrestamoEntity?

    @Query("SELECT * FROM prestamos ORDER BY idPrestamo")
    suspend fun listarTodosInterno(): List<PrestamoEntity>

    @Insert(onConflict = OnConflictStrategy.REPLACE)
    suspend fun insertarTodos(prestamos: List<PrestamoEntity>)

    @Query("DELETE FROM prestamos")
    suspend fun eliminarTodos()

    @Query("DELETE FROM prestamos WHERE idPrestamo = :idPrestamo")
    suspend fun eliminarPorId(idPrestamo: Long)

    @Query("SELECT COUNT(*) FROM prestamos WHERE idCliente = :idCliente")
    suspend fun contarPorCliente(idCliente: Long): Int
}
