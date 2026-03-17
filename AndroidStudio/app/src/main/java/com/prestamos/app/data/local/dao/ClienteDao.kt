package com.prestamos.app.data.local.dao

import androidx.room.Dao
import androidx.room.Insert
import androidx.room.OnConflictStrategy
import androidx.room.Query
import androidx.room.Update
import com.prestamos.app.data.local.entity.ClienteEntity
import kotlinx.coroutines.flow.Flow

@Dao
interface ClienteDao {
    @Insert(onConflict = OnConflictStrategy.ABORT)
    suspend fun insertar(cliente: ClienteEntity): Long

    @Update
    suspend fun actualizar(cliente: ClienteEntity)

    @Query("SELECT * FROM clientes ORDER BY nombre, apellido")
    fun listar(): Flow<List<ClienteEntity>>

    @Query("SELECT * FROM clientes WHERE idCliente = :idCliente LIMIT 1")
    suspend fun obtenerPorId(idCliente: Long): ClienteEntity?

    @Query("SELECT * FROM clientes ORDER BY idCliente")
    suspend fun listarTodosInterno(): List<ClienteEntity>

    @Insert(onConflict = OnConflictStrategy.REPLACE)
    suspend fun insertarTodos(clientes: List<ClienteEntity>)

    @Query("DELETE FROM clientes")
    suspend fun eliminarTodos()

    @Query("SELECT COUNT(*) FROM clientes WHERE documentoIdentidad = :documento")
    suspend fun contarPorDocumento(documento: String): Int

    @Query("DELETE FROM clientes WHERE idCliente = :idCliente")
    suspend fun eliminarPorId(idCliente: Long)
}
