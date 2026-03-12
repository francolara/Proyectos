package com.prestamos.app.data.repository

import androidx.room.withTransaction
import com.prestamos.app.data.local.AppDatabase
import com.prestamos.app.data.local.entity.ClienteEntity
import com.prestamos.app.data.local.entity.CuotaEntity
import com.prestamos.app.data.local.entity.EstadoCuota
import com.prestamos.app.data.local.entity.EstadoPrestamo
import com.prestamos.app.data.local.entity.PagoEntity
import com.prestamos.app.data.local.entity.PrestamoEntity
import com.prestamos.app.data.local.entity.TipoPago
import kotlinx.coroutines.flow.Flow
import java.time.Instant
import java.time.LocalDate
import java.time.ZoneId

class PrestamosRepository(
    private val database: AppDatabase
) {
    private val clienteDao = database.clienteDao()
    private val prestamoDao = database.prestamoDao()
    private val cuotaDao = database.cuotaDao()
    private val pagoDao = database.pagoDao()

    fun observarClientes(): Flow<List<ClienteEntity>> = clienteDao.listar()
    fun observarPrestamos(): Flow<List<PrestamoEntity>> = prestamoDao.listarTodos()
    fun observarCuotasVencidas(fechaActual: Long): Flow<List<CuotaEntity>> = cuotaDao.listarVencidas(fechaActual)
    fun observarTotalCobrado(): Flow<Double?> = pagoDao.totalCobrado()

    fun observarPrestamosPorCliente(idCliente: Long): Flow<List<PrestamoEntity>> = prestamoDao.listarPorCliente(idCliente)
    fun observarCuotasPorPrestamo(idPrestamo: Long): Flow<List<CuotaEntity>> = cuotaDao.listarPorPrestamo(idPrestamo)

    suspend fun registrarCliente(nombre: String, apellido: String, documento: String, nacionalidad: String) {
        val ahora = System.currentTimeMillis()
        clienteDao.insertar(
            ClienteEntity(
                nombre = nombre,
                apellido = apellido,
                documentoIdentidad = documento,
                nacionalidad = nacionalidad,
                fechaRegistro = ahora,
                fechaModificacion = ahora
            )
        )
    }

    suspend fun registrarPrestamo(
        idCliente: Long,
        monto: Double,
        interesPorcentaje: Double,
        tipoPago: TipoPago,
        cantidadCuotas: Int,
        fechaPrimeraCuota: Long
    ) {
        val ahora = System.currentTimeMillis()
        val montoTotal = monto + (monto * (interesPorcentaje / 100.0))
        val montoCuota = montoTotal / cantidadCuotas

        database.withTransaction {
            val idPrestamo = prestamoDao.insertar(
                PrestamoEntity(
                    idCliente = idCliente,
                    montoPrestado = monto,
                    interes = interesPorcentaje,
                    tipoPago = tipoPago,
                    cantidadCuotas = cantidadCuotas,
                    fechaPrimeraCuota = fechaPrimeraCuota,
                    montoTotalPrestamo = montoTotal,
                    montoCuota = montoCuota,
                    estadoPrestamo = EstadoPrestamo.ACTIVO,
                    fechaRegistro = ahora,
                    fechaModificacion = ahora
                )
            )

            val primeraFecha = millisToLocalDate(fechaPrimeraCuota)
            val cuotas = (1..cantidadCuotas).map { numero ->
                val fechaCuota = when (tipoPago) {
                    TipoPago.DIARIO -> primeraFecha.plusDays((numero - 1).toLong())
                    TipoPago.SEMANAL -> primeraFecha.plusWeeks((numero - 1).toLong())
                    TipoPago.MENSUAL -> primeraFecha.plusMonths((numero - 1).toLong())
                }
                CuotaEntity(
                    idPrestamo = idPrestamo,
                    numeroCuota = numero,
                    fechaVencimiento = localDateToMillis(fechaCuota),
                    montoCuota = montoCuota,
                    montoPagado = 0.0,
                    saldoPendiente = montoCuota,
                    estadoCuota = EstadoCuota.PENDIENTE,
                    fechaRegistro = ahora,
                    fechaModificacion = ahora
                )
            }
            cuotaDao.insertarCuotas(cuotas)
        }
    }

    suspend fun registrarPago(
        idPrestamo: Long,
        idCuota: Long,
        montoAbono: Double,
        observacion: String?
    ) {
        require(montoAbono > 0.0) { "El abono debe ser mayor a 0" }

        database.withTransaction {
            val cuota = cuotaDao.obtenerPorId(idCuota)
                ?: error("Cuota no encontrada")
            require(cuota.idPrestamo == idPrestamo) { "La cuota no pertenece al préstamo seleccionado" }
            require(cuota.estadoCuota != EstadoCuota.PAGADO) { "La cuota ya está pagada" }
            require(montoAbono <= cuota.saldoPendiente) { "El abono no puede exceder el saldo pendiente" }

            val ahora = System.currentTimeMillis()
            val nuevoMontoPagado = cuota.montoPagado + montoAbono
            val nuevoSaldo = cuota.montoCuota - nuevoMontoPagado
            val nuevoEstado = when {
                nuevoSaldo <= 0.0 -> EstadoCuota.PAGADO
                nuevoMontoPagado > 0.0 -> EstadoCuota.PARCIAL
                else -> EstadoCuota.PENDIENTE
            }

            pagoDao.insertar(
                PagoEntity(
                    idPrestamo = idPrestamo,
                    idCuota = idCuota,
                    fechaPago = ahora,
                    montoAbono = montoAbono,
                    observacion = observacion?.takeIf { it.isNotBlank() },
                    fechaRegistro = ahora,
                    fechaModificacion = ahora
                )
            )

            cuotaDao.actualizar(
                cuota.copy(
                    montoPagado = nuevoMontoPagado,
                    saldoPendiente = nuevoSaldo.coerceAtLeast(0.0),
                    estadoCuota = nuevoEstado,
                    fechaModificacion = ahora
                )
            )

            val cuotasPrestamo = cuotaDao.listarPorPrestamoInterno(idPrestamo)
            val prestamo = prestamoDao.obtenerPorId(idPrestamo) ?: return@withTransaction
            val estadoPrestamo = if (cuotasPrestamo.all { it.saldoPendiente <= 0.0 }) {
                EstadoPrestamo.PAGADO
            } else {
                EstadoPrestamo.ACTIVO
            }
            prestamoDao.actualizar(prestamo.copy(estadoPrestamo = estadoPrestamo, fechaModificacion = ahora))
        }
    }

    private fun millisToLocalDate(millis: Long): LocalDate =
        Instant.ofEpochMilli(millis).atZone(ZoneId.systemDefault()).toLocalDate()

    private fun localDateToMillis(localDate: LocalDate): Long =
        localDate.atStartOfDay(ZoneId.systemDefault()).toInstant().toEpochMilli()
}
