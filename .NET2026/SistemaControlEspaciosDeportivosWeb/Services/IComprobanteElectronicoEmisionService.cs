namespace SistemaControlEspaciosDeportivosWeb.Services;

public interface IComprobanteElectronicoEmisionService
{
    Task<ComprobanteEmisionResultado> EmitirAsync(int negocioId, int comprobanteId, string usuario);
    Task<ComprobanteEmisionResultado> EmitirManualAsync(int negocioId, int comprobanteId, string usuario);
    Task<ComprobanteEmisionResultado> ConsultarEstadoAsync(int negocioId, int comprobanteId, string usuario);
}

public sealed class ComprobanteEmisionResultado
{
    public bool Exito { get; init; }
    public string Codigo { get; init; } = string.Empty;
    public string Mensaje { get; init; } = string.Empty;
}
