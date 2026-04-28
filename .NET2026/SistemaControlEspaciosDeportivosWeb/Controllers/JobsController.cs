using Microsoft.AspNetCore.Authorization;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Extensions.Options;
using SistemaControlEspaciosDeportivosWeb.Services;
using System.Diagnostics;
using System.Security.Cryptography;
using System.Text;

namespace SistemaControlEspaciosDeportivosWeb.Controllers;

[ApiController]
[AllowAnonymous]
[IgnoreAntiforgeryToken]
[Route("jobs/reservas")]
public class JobsController(
    ISportCenterStoredProcedureService spService,
    IOptions<JobsSettings> jobsOptions,
    IWebHostEnvironment environment,
    ILogger<JobsController> logger) : ControllerBase
{
    private readonly JobsSettings jobs = jobsOptions.Value;

    [HttpPost("autocancelar-no-confirmadas")]
    public Task<IActionResult> AutoCancelarNoConfirmadasPost([FromQuery] string? token = null)
        => EjecutarAutoCancelacionAsync(token);

    [HttpGet("autocancelar-no-confirmadas")]
    public Task<IActionResult> AutoCancelarNoConfirmadasGet([FromQuery] string? token = null)
        => EjecutarAutoCancelacionAsync(token);

    private async Task<IActionResult> EjecutarAutoCancelacionAsync(string? tokenQuery)
    {
        var stopwatch = Stopwatch.StartNew();
        var endpoint = $"{Request.Method} {Request.Path}";

        if (!jobs.AutoCancelEnabled)
        {
            logger.LogWarning("Job bloqueado por configuracion. Endpoint={Endpoint}", endpoint);
            return StatusCode(StatusCodes.Status503ServiceUnavailable, new { ok = false, error = "Job deshabilitado por configuracion." });
        }

        if (!jobs.IsEnvironmentAllowed(environment.EnvironmentName))
        {
            logger.LogWarning("Job bloqueado por entorno. Endpoint={Endpoint}, Entorno={Entorno}", endpoint, environment.EnvironmentName);
            return StatusCode(StatusCodes.Status403Forbidden, new { ok = false, error = "Entorno no permitido para ejecutar el job." });
        }

        if (string.IsNullOrWhiteSpace(jobs.Token))
        {
            logger.LogError("Job sin token configurado. Endpoint={Endpoint}", endpoint);
            return StatusCode(StatusCodes.Status500InternalServerError, new { ok = false, error = "Token de job no configurado." });
        }

        var tokenHeader = Request.Headers["X-Jobs-Token"].FirstOrDefault();
        var tokenRecibido = string.IsNullOrWhiteSpace(tokenHeader) ? tokenQuery : tokenHeader;
        if (!TokenValido(tokenRecibido, jobs.Token))
        {
            logger.LogWarning("Job token invalido. Endpoint={Endpoint}", endpoint);
            return Unauthorized(new { ok = false, error = "Token invalido." });
        }

        try
        {
            logger.LogInformation("Job iniciado. Endpoint={Endpoint}, UsuarioSistema={UsuarioSistema}", endpoint, jobs.UsuarioSistema);

            var procesadas = await spService.ReservasAutoCancelarNoConfirmadasAsync(
                DateTime.Now,
                string.IsNullOrWhiteSpace(jobs.UsuarioSistema) ? "job_scheduler" : jobs.UsuarioSistema);

            stopwatch.Stop();
            logger.LogInformation("Job finalizado OK. Endpoint={Endpoint}, Procesadas={Procesadas}, DuracionMs={DuracionMs}",
                endpoint,
                procesadas,
                stopwatch.ElapsedMilliseconds);

            return Ok(new
            {
                ok = true,
                procesadas,
                duracionMs = stopwatch.ElapsedMilliseconds
            });
        }
        catch (Exception ex)
        {
            stopwatch.Stop();
            logger.LogError(ex, "Job con error. Endpoint={Endpoint}, DuracionMs={DuracionMs}", endpoint, stopwatch.ElapsedMilliseconds);
            return StatusCode(StatusCodes.Status500InternalServerError, new
            {
                ok = false,
                error = "Error ejecutando autocancelacion de reservas.",
                detalle = ex.Message
            });
        }
    }

    private static bool TokenValido(string? recibido, string configurado)
    {
        if (string.IsNullOrWhiteSpace(recibido))
            return false;

        var recibidoBytes = Encoding.UTF8.GetBytes(recibido.Trim());
        var configuradoBytes = Encoding.UTF8.GetBytes(configurado.Trim());
        return recibidoBytes.Length == configuradoBytes.Length &&
               CryptographicOperations.FixedTimeEquals(recibidoBytes, configuradoBytes);
    }
}

