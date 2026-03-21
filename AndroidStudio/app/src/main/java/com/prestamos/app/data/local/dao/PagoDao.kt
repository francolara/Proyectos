package com.prestamos.app.data.local.dao

import androidx.room.Dao
import androidx.room.Insert
import androidx.room.OnConflictStrategy
import androidx.room.Query
import com.prestamos.app.data.local.entity.PagoEntity
import kotlinx.coroutines.flow.Flow

@Dao
interface PagoDao {
    @Insert(onConflict = OnConflictStrategy.ABORT)
    suspend fun insertar(pago: PagoEntity): Long

    @Query("SELECT * FROM pagos WHERE idCuota = :idCuota ORDER BY fechaPago DESC")
    fun listarPorCuota(idCuota: Long): Flow<List<PagoEntity>>

    @Query("SELECT SUM(montoAbono) FROM pagos")
    fun totalCobrado(): Flow<Double?>

    @Query("SELECT * FROM pagos ORDER BY fechaPago DESC")
    fun listarTodos(): Flow<List<PagoEntity>>

    @Query("SELECT * FROM pagos ORDER BY idPago")
    suspend fun listarTodosInterno(): List<PagoEntity>

    @Insert(onConflict = OnConflictStrategy.REPLACE)
    suspend fun insertarTodos(pagos: List<PagoEntity>)

    @Query("DELETE FROM pagos")
    suspend fun eliminarTodos()

    @Query("SELECT COUNT(*) FROM pagos WHERE idPrestamo = :idPrestamo")
    suspend fun contarPorPrestamo(idPrestamo: Long): Int
}
