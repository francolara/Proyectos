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

    @Query("SELECT * FROM pagos ORDER BY fechaPago DESC, idPago DESC")
    fun listarTodos(): Flow<List<PagoEntity>>

    @Query("SELECT * FROM pagos ORDER BY idPago")
    suspend fun listarTodosInterno(): List<PagoEntity>

    @Insert(onConflict = OnConflictStrategy.REPLACE)
    suspend fun insertarTodos(pagos: List<PagoEntity>)

    @Query("DELETE FROM pagos")
    suspend fun eliminarTodos()

    @Query("SELECT COUNT(*) FROM pagos WHERE idPrestamo = :idPrestamo")
    suspend fun contarPorPrestamo(idPrestamo: Long): Int

    @Query("SELECT * FROM pagos WHERE idPago = :idPago LIMIT 1")
    suspend fun obtenerPorId(idPago: Long): PagoEntity?

    @Query("SELECT * FROM pagos WHERE idPrestamo = :idPrestamo ORDER BY fechaPago DESC, idPago DESC LIMIT 1")
    suspend fun obtenerUltimoPorPrestamo(idPrestamo: Long): PagoEntity?

    @Query("DELETE FROM pagos WHERE idPago = :idPago")
    suspend fun eliminarPorId(idPago: Long)
}
