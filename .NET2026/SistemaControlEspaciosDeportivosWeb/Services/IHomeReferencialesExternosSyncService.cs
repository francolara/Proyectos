using SistemaControlEspaciosDeportivosWeb.ViewModels;

namespace SistemaControlEspaciosDeportivosWeb.Services;

public interface IHomeReferencialesExternosSyncService
{
    Task<ReferencialesExternosSyncResultadoViewModel> EjecutarBarridoAsync(
        string codigoUbigeo,
        int tipoDeporteSuperId,
        string palabraClave,
        int maxResultados,
        bool descargarTelefonos,
        bool descargarFotos,
        string usuario,
        CancellationToken cancellationToken = default);
}
