package com.prestamos.app.data.repository

import androidx.room.withTransaction
import com.prestamos.app.data.local.AppDatabase
import com.prestamos.app.data.local.entity.ClienteEntity
import com.prestamos.app.data.local.entity.CuotaEntity
import com.prestamos.app.data.local.entity.EstadoCuota
import com.prestamos.app.data.local.entity.EstadoPrestamo
import com.prestamos.app.data.local.entity.Moneda
import com.prestamos.app.data.local.entity.PagoEntity
import com.prestamos.app.data.local.entity.PrestamoEntity
import com.prestamos.app.data.local.entity.TipoPago
import com.prestamos.app.data.local.entity.TipoCobroEntity
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
    private val tipoCobroDao = database.tipoCobroDao()

    fun observarClientes(): Flow<List<ClienteEntity>> = clienteDao.listar()
    fun observarPrestamos(): Flow<List<PrestamoEntity>> = prestamoDao.listarTodos()
    fun observarCuotas(): Flow<List<CuotaEntity>> = cuotaDao.listarTodas()
    fun observarPagos(): Flow<List<PagoEntity>> = pagoDao.listarTodos()
    fun observarCuotasVencidas(fechaActual: Long): Flow<List<CuotaEntity>> = cuotaDao.listarVencidas(fechaActual)
    fun observarTotalCobrado(): Flow<Double?> = pagoDao.totalCobrado()
    fun observarPrestamosPorCliente(idCliente: Long): Flow<List<PrestamoEntity>> = prestamoDao.listarPorCliente(idCliente)
    fun observarCuotasPorPrestamo(idPrestamo: Long): Flow<List<CuotaEntity>> = cuotaDao.listarPorPrestamo(idPrestamo)
    fun observarTiposCobro(): Flow<List<TipoCobroEntity>> = tipoCobroDao.listar()

    suspend fun registrarCliente(
        nombre: String,
        apellido: String,
        documento: String,
        direccion: String,
        telefono: String
    ) {
        val ahora = System.currentTimeMillis()
        require(clienteDao.contarPorDocumento(documento) == 0) { "Ya existe un cliente con ese documento" }
        clienteDao.insertar(
            ClienteEntity(
                nombre = nombre,
                apellido = apellido,
                documentoIdentidad = documento,
                direccion = direccion,
                telefono = telefono,
                fechaRegistro = ahora,
                fechaModificacion = ahora
            )
        )
    }

    suspend fun actualizarCliente(
        idCliente: Long,
        nombre: String,
        apellido: String,
        direccion: String,
        telefono: String
    ) {
        val ahora = System.currentTimeMillis()
        val cliente = clienteDao.obtenerPorId(idCliente) ?: error("Cliente no encontrado")
        require(nombre.isNotBlank()) { "Nombres obligatorio" }
        require(apellido.isNotBlank()) { "Apellido obligatorio" }
        clienteDao.actualizar(
            cliente.copy(
                nombre = nombre,
                apellido = apellido,
                direccion = direccion,
                telefono = telefono,
                fechaModificacion = ahora
            )
        )
    }

    suspend fun eliminarClienteSiNoTienePrestamos(idCliente: Long) {
        database.withTransaction {
            val cliente = clienteDao.obtenerPorId(idCliente) ?: error("Cliente no encontrado")
            val totalPrestamos = prestamoDao.contarPorCliente(idCliente)
            require(totalPrestamos == 0) { "No se puede eliminar: el cliente tiene prestamos registrados" }
            clienteDao.eliminarPorId(cliente.idCliente)
        }
    }

    suspend fun registrarPrestamo(
        idCliente: Long,
        monto: Double,
        interesPorcentaje: Double,
        moneda: Moneda,
        tipoPago: TipoPago,
        intervaloDiasPersonalizado: Int?,
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
                    moneda = moneda,
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
            val diasPersonalizado = if (tipoPago == TipoPago.PERSONALIZADO) {
                require((intervaloDiasPersonalizado ?: 0) > 0) { "Intervalo personalizado invalido" }
                intervaloDiasPersonalizado
            } else {
                null
            }
            val cuotas = (1..cantidadCuotas).map { numero ->
                val fechaCuota = when (tipoPago) {
                    TipoPago.DIARIO -> primeraFecha.plusDays((numero - 1).toLong())
                    TipoPago.SEMANAL -> primeraFecha.plusWeeks((numero - 1).toLong())
                    TipoPago.QUINCENAL -> primeraFecha.plusDays((numero - 1).toLong() * 15L)
                    TipoPago.MENSUAL -> primeraFecha.plusMonths((numero - 1).toLong())
                    TipoPago.PERSONALIZADO -> primeraFecha.plusDays((numero - 1).toLong() * diasPersonalizado!!.toLong())
                }
                CuotaEntity(
                    idPrestamo = idPrestamo,
                    numeroCuota = numero,
                    fechaVencimiento = localDateToMillis(fechaCuota),
                    montoCuota = montoCuota,
                    montoPagado = 0.0,
                    moraPendiente = 0.0,
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
        idTipoCobro: Long,
        observacion: String?
    ) {
        require(montoAbono > 0.0) { "El abono debe ser mayor a 0" }

        database.withTransaction {
            val cuota = cuotaDao.obtenerPorId(idCuota) ?: error("Cuota no encontrada")
            require(cuota.idPrestamo == idPrestamo) { "La cuota no pertenece al prestamo seleccionado" }
            require(cuota.estadoCuota != EstadoCuota.PAGADO) { "La cuota ya esta pagada" }
            require(montoAbono <= cuota.saldoPendiente) { "El abono no puede exceder el saldo pendiente" }

            val siguienteCuotaPendiente = cuotaDao.listarPorPrestamoInterno(idPrestamo)
                .firstOrNull { it.saldoPendiente > 0.0 }
            require(siguienteCuotaPendiente != null) { "El prestamo no tiene cuotas pendientes" }
            require(cuota.idCuota == siguienteCuotaPendiente.idCuota) {
                "Debe registrar primero la cuota ${siguienteCuotaPendiente.numeroCuota}"
            }
            val tipoCobro = tipoCobroDao.obtenerPorId(idTipoCobro) ?: error("Tipo de cobro no encontrado")

            val ahora = System.currentTimeMillis()
            val capitalPendiente = (cuota.montoCuota - cuota.montoPagado).coerceAtLeast(0.0)
            val moraCobrada = montoAbono.coerceAtMost(cuota.moraPendiente)
            val abonoCapital = (montoAbono - moraCobrada).coerceAtLeast(0.0)
            val nuevoMontoPagado = (cuota.montoPagado + abonoCapital).coerceAtMost(cuota.montoCuota)
            val nuevaMoraPendiente = (cuota.moraPendiente - moraCobrada).coerceAtLeast(0.0)
            val nuevoCapitalPendiente = (capitalPendiente - abonoCapital).coerceAtLeast(0.0)
            val nuevoSaldo = nuevoCapitalPendiente + nuevaMoraPendiente
            val nuevoEstado = when {
                nuevoSaldo <= 0.0 -> EstadoCuota.PAGADO
                nuevoMontoPagado > 0.0 -> EstadoCuota.PARCIAL
                else -> EstadoCuota.PENDIENTE
            }

            pagoDao.insertar(
                PagoEntity(
                    idPrestamo = idPrestamo,
                    idCuota = idCuota,
                    idTipoCobro = tipoCobro.idTipoCobro,
                    fechaPago = ahora,
                    montoAbono = montoAbono,
                    moraCobrada = moraCobrada,
                    observacion = observacion?.takeIf { it.isNotBlank() },
                    fechaRegistro = ahora,
                    fechaModificacion = ahora
                )
            )

            cuotaDao.actualizar(
                cuota.copy(
                    montoPagado = nuevoMontoPagado,
                    moraPendiente = nuevaMoraPendiente,
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

    suspend fun registrarTipoCobro(nombre: String) {
        val nombreLimpio = nombre.trim()
        require(nombreLimpio.isNotBlank()) { "Ingresa un tipo de cobro valido" }
        require(tipoCobroDao.contarPorNombre(nombreLimpio) == 0) { "Ese tipo de cobro ya existe" }
        val ahora = System.currentTimeMillis()
        tipoCobroDao.insertar(
            TipoCobroEntity(
                nombre = nombreLimpio,
                fechaRegistro = ahora,
                fechaModificacion = ahora
            )
        )
    }

    suspend fun eliminarTipoCobroSiNoTienePagos(idTipoCobro: Long) {
        database.withTransaction {
            val tipo = tipoCobroDao.obtenerPorId(idTipoCobro) ?: error("Tipo de cobro no encontrado")
            val usos = pagoDao.contarPorTipoCobro(idTipoCobro)
            require(usos == 0) { "No se puede eliminar: ya fue usado en cobros" }
            tipoCobroDao.eliminarPorId(tipo.idTipoCobro)
        }
    }

    suspend fun eliminarPrestamoSiNoTienePagos(idPrestamo: Long) {
        database.withTransaction {
            val prestamo = prestamoDao.obtenerPorId(idPrestamo) ?: error("Prestamo no encontrado")
            val totalPagos = pagoDao.contarPorPrestamo(idPrestamo)
            require(totalPagos == 0) { "No se puede eliminar: el prestamo ya tiene pagos registrados" }
            prestamoDao.eliminarPorId(prestamo.idPrestamo)
        }
    }

    suspend fun eliminarPagoSiEsUltimo(idPago: Long) {
        database.withTransaction {
            val pago = pagoDao.obtenerPorId(idPago) ?: error("Pago no encontrado")
            val ultimoPago = pagoDao.obtenerUltimoPorPrestamo(pago.idPrestamo)
                ?: error("No hay pagos para este prestamo")
            require(ultimoPago.idPago == pago.idPago) {
                "Solo se puede eliminar el ultimo pago del prestamo"
            }

            val cuota = cuotaDao.obtenerPorId(pago.idCuota) ?: error("Cuota no encontrada")
            val ahora = System.currentTimeMillis()
            val abonoCapitalRevertido = (pago.montoAbono - pago.moraCobrada).coerceAtLeast(0.0)
            val nuevoMontoPagado = (cuota.montoPagado - abonoCapitalRevertido).coerceAtLeast(0.0)
            val nuevaMoraPendiente = (cuota.moraPendiente + pago.moraCobrada).coerceAtLeast(0.0)
            val nuevoSaldo = (cuota.montoCuota - nuevoMontoPagado).coerceAtLeast(0.0) + nuevaMoraPendiente
            val nuevoEstado = when {
                nuevoMontoPagado <= 0.0 -> EstadoCuota.PENDIENTE
                nuevoSaldo <= 0.0 -> EstadoCuota.PAGADO
                else -> EstadoCuota.PARCIAL
            }

            cuotaDao.actualizar(
                cuota.copy(
                    montoPagado = nuevoMontoPagado,
                    moraPendiente = nuevaMoraPendiente,
                    saldoPendiente = nuevoSaldo,
                    estadoCuota = nuevoEstado,
                    fechaModificacion = ahora
                )
            )
            pagoDao.eliminarPorId(pago.idPago)

            val cuotasPrestamo = cuotaDao.listarPorPrestamoInterno(pago.idPrestamo)
            val prestamo = prestamoDao.obtenerPorId(pago.idPrestamo) ?: return@withTransaction
            val estadoPrestamo = if (cuotasPrestamo.all { it.saldoPendiente <= 0.0 }) {
                EstadoPrestamo.PAGADO
            } else {
                EstadoPrestamo.ACTIVO
            }
            prestamoDao.actualizar(prestamo.copy(estadoPrestamo = estadoPrestamo, fechaModificacion = ahora))
        }
    }

    suspend fun aplicarMoraManualCuotaVencida(
        idCuota: Long,
        montoMora: Double
    ) {
        require(montoMora > 0.0) { "La mora debe ser mayor a 0" }
        database.withTransaction {
            val cuota = cuotaDao.obtenerPorId(idCuota) ?: error("Cuota no encontrada")
            require(cuota.saldoPendiente > 0.0) { "La cuota no tiene saldo pendiente" }
            val hoy = System.currentTimeMillis()
            require(cuota.fechaVencimiento < hoy) { "La cuota aun no esta vencida" }

            val ahora = System.currentTimeMillis()
            val nuevaMora = cuota.moraPendiente + montoMora
            val nuevoSaldo = cuota.saldoPendiente + montoMora
            cuotaDao.actualizar(
                cuota.copy(
                    moraPendiente = nuevaMora,
                    saldoPendiente = nuevoSaldo,
                    fechaModificacion = ahora
                )
            )

        }
    }

    private fun millisToLocalDate(millis: Long): LocalDate =
        Instant.ofEpochMilli(millis).atZone(ZoneId.systemDefault()).toLocalDate()

    private fun localDateToMillis(localDate: LocalDate): Long =
        localDate.atStartOfDay(ZoneId.systemDefault()).toInstant().toEpochMilli()
}
