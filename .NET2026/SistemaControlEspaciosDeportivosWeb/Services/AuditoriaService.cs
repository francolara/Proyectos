using System.Security.Claims;
using System.Text.Json;
using Microsoft.EntityFrameworkCore;
using SistemaControlEspaciosDeportivosWeb.Data;
using SistemaControlEspaciosDeportivosWeb.Models;

namespace SistemaControlEspaciosDeportivosWeb.Services;

public interface IAuditoriaService
{
    Task RegistrarAsync(string modulo, string accion, string entidad, string entidadId, int? negocioId, object? detalle = null);
}

public class AuditoriaService(
    ApplicationDbContext dbContext,
    IHttpContextAccessor httpContextAccessor) : IAuditoriaService
{
    public async Task RegistrarAsync(string modulo, string accion, string entidad, string entidadId, int? negocioId, object? detalle = null)
    {
        var user = httpContextAccessor.HttpContext?.User;
        var usuarioId = user?.FindFirstValue(ClaimTypes.NameIdentifier) ?? "sistema";
        var usuarioNombre = user?.Identity?.Name;
        var usuarioCorreo = user?.FindFirstValue(ClaimTypes.Email);

        var registro = new BitacoraAuditoria
        {
            NegocioId = negocioId,
            Modulo = modulo,
            Accion = accion,
            Entidad = entidad,
            EntidadId = entidadId,
            UsuarioId = usuarioId,
            UsuarioNombre = usuarioNombre,
            UsuarioCorreo = usuarioCorreo,
            DetalleJson = detalle is null ? null : JsonSerializer.Serialize(detalle),
            FechaRegistro = DateTime.UtcNow
        };

        dbContext.BitacoraAuditoria.Add(registro);
        await dbContext.SaveChangesAsync();
    }
}

