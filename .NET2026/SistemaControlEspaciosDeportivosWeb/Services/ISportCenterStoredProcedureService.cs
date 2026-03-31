using Microsoft.AspNetCore.Mvc.Rendering;
using SistemaControlEspaciosDeportivosWeb.ViewModels;

namespace SistemaControlEspaciosDeportivosWeb.Services;

public interface ISportCenterStoredProcedureService
{
    Task<List<SedePublicaViewModel>> HomeListarSedesAsync();
    Task<List<TipoDeportePublicoViewModel>> HomeListarTiposDeporteAsync();
    Task<List<EspacioDisponibleViewModel>> HomeBuscarEspaciosDisponiblesAsync(DateOnly fecha, TimeOnly horaInicio, TimeOnly horaFin, int? sedeId, int? tipoDeporteId);
    Task<string> HomeSolicitarReservaPublicaAsync(SolicitudReservaPublicaFormViewModel model);
    Task<SolicitudPublicaDetalleViewModel?> HomeConsultarSolicitudAsync(string codigoSolicitud, string telefono);
    Task<SolicitudNotificacionEmailViewModel?> HomeObtenerSolicitudParaNotificacionAsync(string codigoSolicitud);
    Task<bool> HomeMarcarSolicitudNotificadaAsync(string codigoSolicitud);
    Task<ConfiguracionClubViewModel?> ConfiguracionClubObtenerAsync(int negocioId);
    Task<bool> ConfiguracionClubActualizarAsync(ConfiguracionClubViewModel model, string usuario);
    Task<List<SelectListItem>> ConfiguracionClubComboMonedasAsync();

    Task<List<NegocioAccesoViewModel>> PanelListarNegociosUsuarioAsync(string usuarioId);
    Task<string?> PanelObtenerRolAsync(string usuarioId, int negocioId);
    Task<List<PermisoModuloViewModel>> PanelListarModulosPermitidosAsync(string usuarioId, int negocioId);
    Task<(int TotalSedes, int TotalEspacios, int ReservasHoy, decimal IngresosHoy, decimal OcupacionHoyPct, int NoShowMes, decimal TicketPromedioMes)> PanelObtenerMetricasAsync(int negocioId, DateOnly fecha, int? sedeId = null);

    Task<List<SedeItemViewModel>> SedesListarAsync(int negocioId, int? sedeId = null);
    Task<SedeFormViewModel?> SedesObtenerAsync(int negocioId, int id);
    Task<int> SedesCrearAsync(SedeFormViewModel model, string usuario);
    Task<bool> SedesActualizarAsync(SedeFormViewModel model, string usuario);
    Task<bool> SedesEliminarAsync(int negocioId, int id, string usuario);
    Task<List<SelectListItem>> SedesComboServiciosAsync();

    Task<List<EspacioItemViewModel>> EspaciosListarAsync(int negocioId, int? sedeId = null);
    Task<EspacioFormViewModel?> EspaciosObtenerAsync(int negocioId, int id);
    Task<int> EspaciosCrearAsync(EspacioFormViewModel model, string usuario);
    Task<bool> EspaciosActualizarAsync(EspacioFormViewModel model, string usuario);
    Task<bool> EspaciosEliminarAsync(int negocioId, int id, string usuario);
    Task<List<SelectListItem>> EspaciosComboSedesAsync(int negocioId, int? sedeId = null);
    Task<List<SelectListItem>> EspaciosComboTiposDeporteAsync();
    Task<List<SelectListItem>> EspaciosComboTiposSueloAsync();

    Task<List<ReservaItemViewModel>> ReservasListarAsync(int negocioId, DateOnly? fechaDesde = null, DateOnly? fechaHasta = null, int? sedeId = null, int? espacioDeportivoId = null, int? estado = null, string? estadosCsv = null);
    Task<ReservaFormViewModel?> ReservasObtenerAsync(int negocioId, int id);
    Task<int> ReservasCrearAsync(ReservaFormViewModel model, string usuario);
    Task<bool> ReservasActualizarAsync(ReservaFormViewModel model, string usuario);
    Task<bool> ReservasEliminarAsync(int negocioId, int id, string usuario);
    Task<bool> ReservasCambiarEstadoRapidoAsync(int negocioId, int id, int nuevoEstado, string usuario);
    Task<List<ReservaCalendarioEventoViewModel>> ReservasCalendarioEventosAsync(int negocioId, DateOnly fechaDesde, DateOnly fechaHasta, int? sedeId = null, int? espacioDeportivoId = null, int? estado = null);
    Task<bool> ReservasMoverAsync(int negocioId, int id, DateOnly fecha, TimeOnly horaInicio, TimeOnly horaFin, string usuario);
    Task<ReservaDisponibilidadValidacionViewModel> ReservasValidarDisponibilidadAsync(int negocioId, int? reservaId, int espacioDeportivoId, DateOnly fecha, TimeOnly horaInicio, TimeOnly horaFin);
    Task<List<BloqueoHorarioItemViewModel>> BloqueosListarAsync(int negocioId, DateOnly fechaDesde, DateOnly fechaHasta, int? sedeId = null, int? espacioDeportivoId = null);
    Task<int> BloqueosCrearAsync(BloqueoHorarioFormViewModel model, string usuario);
    Task<bool> BloqueosEliminarAsync(int negocioId, int id, string usuario);
    Task<List<SelectListItem>> ReservasComboEspaciosAsync(int negocioId, int? sedeId = null);
    Task<List<SelectListItem>> ReservasComboClientesAsync(int negocioId);

    Task<List<PagoItemViewModel>> PagosListarAsync(int negocioId, int? sedeId = null);
    Task<PagoFormViewModel?> PagosObtenerAsync(int negocioId, int id);
    Task<int> PagosCrearAsync(PagoFormViewModel model, string usuario);
    Task<bool> PagosActualizarAsync(PagoFormViewModel model, string usuario);
    Task<bool> PagosEliminarAsync(int negocioId, int id, string usuario);
    Task<List<SelectListItem>> PagosComboReservasAsync(int negocioId, int? sedeId = null);

    Task<List<ComprobanteItemViewModel>> ComprobantesListarAsync(int negocioId, int? sedeId = null);
    Task<ComprobanteFormViewModel?> ComprobantesObtenerAsync(int negocioId, int id);
    Task<int> ComprobantesCrearAsync(ComprobanteFormViewModel model, string usuario);
    Task<bool> ComprobantesActualizarAsync(ComprobanteFormViewModel model, string usuario);
    Task<bool> ComprobantesEliminarAsync(int negocioId, int id, string usuario);
    Task<List<SelectListItem>> ComprobantesComboReservasAsync(int negocioId, int? sedeId = null);

    Task<List<ClienteItemViewModel>> ClientesListarAsync(int negocioId);
    Task<ClienteFormViewModel?> ClientesObtenerAsync(int negocioId, int id);
    Task<int> ClientesCrearAsync(ClienteFormViewModel model, string usuario);
    Task<bool> ClientesActualizarAsync(ClienteFormViewModel model, string usuario);
    Task<bool> ClientesEliminarAsync(int negocioId, int id, string usuario);

    Task<List<ReporteOcupacionItemViewModel>> ReportesOcupacionPorEspacioAsync(int negocioId, DateOnly fechaDesde, DateOnly fechaHasta, int? sedeId = null);
    Task<List<ReporteIngresoDiaItemViewModel>> ReportesIngresosPorDiaAsync(int negocioId, DateOnly fechaDesde, DateOnly fechaHasta, int? sedeId = null);

    Task<List<SolicitudPublicaItemViewModel>> SolicitudesPublicasListarAsync(int negocioId, DateOnly? fechaDesde = null, DateOnly? fechaHasta = null, int? estado = null);
    Task<bool> SolicitudesPublicasActualizarEstadoAsync(SolicitudEstadoFormViewModel model, string usuario);
    Task<int> SolicitudesPublicasConvertirAReservaAsync(SolicitudConvertirFormViewModel model, string usuario);
    Task<string> HomeSolicitarAltaClubAsync(AltaClubSolicitudFormViewModel model);
    Task<string> HomeRegistrarClubConPruebaAsync(AltaClubSolicitudFormViewModel model, string usuarioId);
    Task<List<AltaClubItemViewModel>> AltasClubesListarAsync(int? estado = null);
    Task<bool> AltasClubesAprobarAsync(int id, string usuario, string? comentarioGestion = null);
    Task<bool> AltasClubesRechazarAsync(int id, string usuario, string? comentarioGestion = null);

    Task<List<UsuarioNegocioItemViewModel>> UsuariosNegocioListarAsync(int negocioId, int? sedeId = null);
    Task<bool> UsuariosNegocioAsignarPorCorreoAsync(int negocioId, string correo, int rolNegocio, int? sedeId, string usuario);
    Task<bool> UsuariosNegocioActualizarRolAsync(int negocioId, int usuarioNegocioId, int rolNegocio, int? sedeId, string usuario);
    Task<bool> UsuariosNegocioDesactivarAsync(int negocioId, int usuarioNegocioId, string usuario);
    Task<List<UsuarioNegocioPermisoModuloViewModel>> UsuariosNegocioPermisosListarAsync(int negocioId, int usuarioNegocioId);
    Task<bool> UsuariosNegocioPermisoGuardarAsync(int negocioId, int usuarioNegocioId, UsuarioNegocioPermisoModuloViewModel model, string usuario);

    Task<List<PromocionItemViewModel>> PromocionesListarAsync(int negocioId, int? sedeId = null);
    Task<PromocionFormViewModel?> PromocionesObtenerAsync(int negocioId, int id);
    Task<int> PromocionesCrearAsync(PromocionFormViewModel model, string usuario);
    Task<bool> PromocionesActualizarAsync(PromocionFormViewModel model, string usuario);
    Task<bool> PromocionesEliminarAsync(int negocioId, int id, string usuario);

    Task<List<ReservaRecordatorioPendienteViewModel>> ReservasRecordatoriosPendientesAsync(DateTime fechaHoraActual);
    Task<bool> ReservasMarcarRecordatorioEnviadoAsync(int negocioId, int reservaId, string usuario);
    Task<int> ReservasAutoNoShowAsync(DateTime fechaHoraActual, string usuario);
}
