package com.prestamos.app.data.local.dao

import androidx.room.Dao
import androidx.room.Insert
import androidx.room.OnConflictStrategy
import androidx.room.Query
import com.prestamos.app.data.local.entity.TipoCobroEntity
import kotlinx.coroutines.flow.Flow

@Dao
interface TipoCobroDao {
    @Query("SELECT * FROM tipos_cobro ORDER BY nombre COLLATE NOCASE ASC")
    fun listar(): Flow<List<TipoCobroEntity>>

    @Query("SELECT * FROM tipos_cobro ORDER BY nombre COLLATE NOCASE ASC")
    suspend fun listarInterno(): List<TipoCobroEntity>

    @Insert(onConflict = OnConflictStrategy.ABORT)
    suspend fun insertar(tipoCobro: TipoCobroEntity): Long

    @Query("DELETE FROM tipos_cobro WHERE idTipoCobro = :idTipoCobro")
    suspend fun eliminarPorId(idTipoCobro: Long)

    @Query("SELECT COUNT(*) FROM tipos_cobro WHERE LOWER(TRIM(nombre)) = LOWER(TRIM(:nombre))")
    suspend fun contarPorNombre(nombre: String): Int

    @Query("SELECT * FROM tipos_cobro WHERE idTipoCobro = :idTipoCobro LIMIT 1")
    suspend fun obtenerPorId(idTipoCobro: Long): TipoCobroEntity?
}

