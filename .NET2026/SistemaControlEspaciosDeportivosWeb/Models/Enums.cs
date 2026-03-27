namespace SistemaControlEspaciosDeportivosWeb.Models;

public enum EstadoEspacioDeportivo
{
    Activo = 1,
    EnMantenimiento = 2,
    Inactivo = 3
}

public enum EstadoReserva
{
    Pendiente = 1,
    Confirmada = 2,
    EnUso = 3,
    Finalizada = 4,
    Cancelada = 5,
    NoAsistio = 6
}

public enum FormaPago
{
    Efectivo = 1,
    Yape = 2,
    Plin = 3,
    Transferencia = 4,
    Tarjeta = 5
}

public enum RolNegocio
{
    Administrador = 1,
    Trabajador = 2,
    Recepcion = 3,
    Caja = 4,
    Supervisor = 5
}

public enum TipoComprobante
{
    Boleta = 1,
    Factura = 2
}

public enum EstadoComprobanteElectronico
{
    PendienteEnvio = 1,
    EnviadoSunat = 2,
    AceptadoSunat = 3,
    RechazadoSunat = 4,
    Anulado = 5
}

public enum TipoMoneda
{
    PEN = 1,
    USD = 2
}
