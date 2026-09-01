using Microsoft.AspNetCore.Authorization;
using Microsoft.AspNetCore.Mvc;
using SistemaAdministrativoWeb.Infrastructure.Contabilidad;
using SistemaAdministrativoWeb.Infrastructure.Empresas;
using SistemaAdministrativoWeb.Infrastructure.Security;
using SistemaAdministrativoWeb.ViewModels.Contabilidad;

namespace SistemaAdministrativoWeb.Controllers;

[Authorize]
[ModulePermission("CUENTACORRIENTE")]
public class CuentaCorrienteController(
    ICurrentCompanyAccessor currentCompanyAccessor,
    ICuentaCorrienteRepository cuentaCorrienteRepository,
    IBancoRepository bancoRepository,
    IMonedaRepository monedaRepository) : Controller
{
    private const int TamanoPagina = 20;
    private const int TamanoAyudaBancos = 100;

    [HttpGet]
    public async Task<IActionResult> Index(string? textoBusqueda = null, int pagina = 1, CancellationToken cancellationToken = default)
    {
        if (!currentCompanyAccessor.TieneEmpresaActiva || !currentCompanyAccessor.EmpresaId.HasValue)
        {
            return RedirectToAction("Index", "EmpresaContexto");
        }

        ViewData["AdminShell"] = true;

        var resultado = await cuentaCorrienteRepository.ListarPaginadoPorEmpresaAsync(
            currentCompanyAccessor.EmpresaId.Value,
            textoBusqueda,
            pagina,
            TamanoPagina,
            false,
            cancellationToken);
        var monedas = await monedaRepository.ListarActivasAsync(cancellationToken);

        var model = ConstruirViewModel(resultado.Items, null, monedas);
        model.TextoBusqueda = textoBusqueda?.Trim() ?? string.Empty;
        model.TotalCuentasCorrientes = resultado.TotalRecords;
        model.Paginacion = new PaginacionViewModel
        {
            PaginaActual = pagina,
            TamanoPagina = TamanoPagina,
            TotalRegistros = resultado.TotalRecords
        };

        return View(model);
    }

    [HttpGet]
    public async Task<IActionResult> Registrar(CancellationToken cancellationToken = default)
    {
        if (!currentCompanyAccessor.TieneEmpresaActiva || !currentCompanyAccessor.EmpresaId.HasValue)
        {
            return RedirectToAction("Index", "EmpresaContexto");
        }

        ViewData["AdminShell"] = true;
        var monedas = await monedaRepository.ListarActivasAsync(cancellationToken);
        return View("Formulario", ConstruirViewModel([], null, monedas));
    }

    [HttpGet]
    public async Task<IActionResult> Editar(int idBancoConfiguracionEmpresa, CancellationToken cancellationToken = default)
    {
        if (!currentCompanyAccessor.TieneEmpresaActiva || !currentCompanyAccessor.EmpresaId.HasValue)
        {
            return RedirectToAction("Index", "EmpresaContexto");
        }

        ViewData["AdminShell"] = true;

        var cuenta = await cuentaCorrienteRepository.ObtenerPorIdAsync(currentCompanyAccessor.EmpresaId.Value, idBancoConfiguracionEmpresa, cancellationToken);
        var monedas = await monedaRepository.ListarActivasAsync(cancellationToken);
        if (cuenta is null)
        {
            return RedirectToAction(nameof(Index));
        }

        return View("Formulario", ConstruirViewModel([], cuenta, monedas));
    }

    [HttpPost]
    [ValidateAntiForgeryToken]
    [ModulePermission("CUENTACORRIENTE", ModulePermissionOperation.Delete)]
    public async Task<IActionResult> Eliminar(int idBancoConfiguracionEmpresa, string? textoBusqueda = null, int pagina = 1, CancellationToken cancellationToken = default)
    {
        if (!currentCompanyAccessor.TieneEmpresaActiva || !currentCompanyAccessor.EmpresaId.HasValue)
        {
            return RedirectToAction("Index", "EmpresaContexto");
        }

        try
        {
            await cuentaCorrienteRepository.EliminarAsync(currentCompanyAccessor.EmpresaId.Value, idBancoConfiguracionEmpresa, cancellationToken);
            TempData["CuentaCorrienteOk"] = "Cuenta corriente eliminada correctamente.";
        }
        catch (Exception ex)
        {
            TempData["CuentaCorrienteError"] = ex.Message;
        }

        return RedirectToAction(nameof(Index), new { textoBusqueda, pagina });
    }

    [HttpGet]
    public async Task<IActionResult> BuscarBancosAyuda(string? textoBusqueda = null, int tamanoPagina = TamanoAyudaBancos, CancellationToken cancellationToken = default)
    {
        var filtro = string.IsNullOrWhiteSpace(textoBusqueda) ? null : textoBusqueda.Trim();
        if (!string.IsNullOrWhiteSpace(filtro) && filtro.Length < 2)
        {
            filtro = null;
        }

        var resultado = await bancoRepository.ListarPaginadoAsync(
            filtro,
            1,
            Math.Clamp(tamanoPagina, 1, TamanoAyudaBancos),
            true,
            cancellationToken);

        return Json(new
        {
            ok = true,
            items = resultado.Items.Select(x => new
            {
                idBanco = x.IdBanco,
                codigoBanco = x.CodigoBanco,
                nombreBanco = x.NombreBanco
            }),
            total = resultado.TotalRecords,
            limitado = resultado.TotalRecords > resultado.Items.Count
        });
    }

    [HttpPost]
    [ValidateAntiForgeryToken]
    [ModuleSavePermission("CUENTACORRIENTE", nameof(CuentaCorrienteFormViewModel.IdBancoConfiguracionEmpresa))]
    public async Task<IActionResult> Guardar(CuentaCorrienteFormViewModel formulario, CancellationToken cancellationToken)
    {
        if (!currentCompanyAccessor.TieneEmpresaActiva || !currentCompanyAccessor.EmpresaId.HasValue)
        {
            return RedirectToAction("Index", "EmpresaContexto");
        }

        ViewData["AdminShell"] = true;
        var monedas = await monedaRepository.ListarActivasAsync(cancellationToken);
        formulario.Monedas = monedas.ToList();

        if (!ModelState.IsValid)
        {
            return View("Formulario", new CuentaCorrienteIndexViewModel
            {
                IdEmpresa = currentCompanyAccessor.EmpresaId.Value,
                EmpresaNombre = currentCompanyAccessor.EmpresaNombre ?? "Empresa activa",
                Monedas = monedas.ToList(),
                Formulario = formulario
            });
        }

        try
        {
            await cuentaCorrienteRepository.GuardarAsync(new GuardarBancoConfiguracionEmpresaRequest
            {
                IdBancoConfiguracionEmpresa = formulario.IdBancoConfiguracionEmpresa,
                IdEmpresa = currentCompanyAccessor.EmpresaId.Value,
                IdBanco = formulario.IdBanco!.Value,
                NroCuentaCorriente = formulario.NroCuentaCorriente.Trim(),
                Titular = formulario.Titular.Trim(),
                IdMoneda = formulario.IdMoneda!.Value,
                IdPlanCuenta = formulario.IdPlanCuenta!.Value,
                PeriodoSaldoInicial = NormalizarPeriodoSaldoInicial(formulario.PeriodoSaldoInicial),
                SaldoInicialDebe = formulario.SaldoInicialDebe,
                SaldoInicialHaber = formulario.SaldoInicialHaber,
                Activo = formulario.Activo,
                UsuarioRegistro = User.Identity?.Name
            }, cancellationToken);

            TempData["CuentaCorrienteOk"] = formulario.IdBancoConfiguracionEmpresa.HasValue
                ? "Cuenta corriente actualizada correctamente."
                : "Cuenta corriente registrada correctamente.";

            return RedirectToAction(nameof(Index));
        }
        catch (Exception ex)
        {
            ModelState.AddModelError(string.Empty, ex.Message);
            return View("Formulario", new CuentaCorrienteIndexViewModel
            {
                IdEmpresa = currentCompanyAccessor.EmpresaId.Value,
                EmpresaNombre = currentCompanyAccessor.EmpresaNombre ?? "Empresa activa",
                Monedas = monedas.ToList(),
                Formulario = formulario
            });
        }
    }

    private CuentaCorrienteIndexViewModel ConstruirViewModel(IReadOnlyCollection<BancoConfiguracionEmpresaDto> cuentas, BancoConfiguracionEmpresaDto? cuentaEditar, IReadOnlyCollection<MonedaDto> monedas)
    {
        var items = cuentas
            .Select(x => new CuentaCorrienteItemViewModel
            {
                IdBancoConfiguracionEmpresa = x.IdBancoConfiguracionEmpresa,
                CodigoBanco = x.CodigoBanco,
                NombreBanco = x.NombreBanco,
                NroCuentaCorriente = x.NroCuentaCorriente,
                Titular = x.Titular,
                IdMoneda = x.IdMoneda,
                MonedaTexto = string.IsNullOrWhiteSpace(x.CodigoMoneda)
                    ? string.Empty
                    : $"{x.CodigoMoneda} - {x.NombreMoneda}",
                IdPlanCuenta = x.IdPlanCuenta,
                CodigoCuenta = x.CodigoCuenta,
                NombreCuenta = x.NombreCuenta,
                PeriodoSaldoInicial = FormatearPeriodoSaldoInicial(x.PeriodoSaldoInicial),
                SaldoInicialDebe = x.SaldoInicialDebe,
                SaldoInicialHaber = x.SaldoInicialHaber,
                Activo = x.Activo,
                FechaRegistro = x.FechaRegistro,
                UsuarioRegistro = x.UsuarioRegistro
            })
            .ToList();

        return new CuentaCorrienteIndexViewModel
        {
            IdEmpresa = currentCompanyAccessor.EmpresaId ?? 0,
            EmpresaNombre = currentCompanyAccessor.EmpresaNombre ?? "Empresa activa",
            TotalCuentasCorrientes = items.Count,
            TotalActivas = items.Count(x => x.Activo),
            TotalBancos = cuentas.Select(x => x.IdBanco).Distinct().Count(),
            Monedas = monedas.ToList(),
            CuentasCorrientes = items,
            Formulario = cuentaEditar is null
                ? new CuentaCorrienteFormViewModel
                {
                    PeriodoSaldoInicial = DateTime.Today.ToString("yyyy-MM"),
                    Monedas = monedas.ToList()
                }
                : new CuentaCorrienteFormViewModel
                {
                    IdBancoConfiguracionEmpresa = cuentaEditar.IdBancoConfiguracionEmpresa,
                    IdBanco = cuentaEditar.IdBanco,
                    BancoTexto = $"{cuentaEditar.CodigoBanco} - {cuentaEditar.NombreBanco}",
                    NroCuentaCorriente = cuentaEditar.NroCuentaCorriente,
                    Titular = cuentaEditar.Titular,
                    IdMoneda = cuentaEditar.IdMoneda,
                    IdPlanCuenta = cuentaEditar.IdPlanCuenta,
                    CuentaTexto = $"{cuentaEditar.CodigoCuenta} - {cuentaEditar.NombreCuenta}",
                    PeriodoSaldoInicial = FormatearPeriodoSaldoInicial(cuentaEditar.PeriodoSaldoInicial),
                    SaldoInicialDebe = cuentaEditar.SaldoInicialDebe,
                    SaldoInicialHaber = cuentaEditar.SaldoInicialHaber,
                    Activo = cuentaEditar.Activo,
                    FechaRegistro = cuentaEditar.FechaRegistro,
                    UsuarioRegistro = cuentaEditar.UsuarioRegistro,
                    Monedas = monedas.ToList()
                }
        };
    }

    private static string FormatearPeriodoSaldoInicial(string? periodo)
    {
        if (!string.IsNullOrWhiteSpace(periodo)
            && periodo.Length == 6
            && int.TryParse(periodo[..4], out _)
            && int.TryParse(periodo[4..], out var mes)
            && mes is >= 1 and <= 12)
        {
            return $"{periodo[..4]}-{periodo[4..]}";
        }

        return DateTime.Today.ToString("yyyy-MM");
    }

    private static string NormalizarPeriodoSaldoInicial(string periodo)
    {
        var periodoTrabajo = (periodo ?? string.Empty).Trim();
        return periodoTrabajo.Length == 7 && periodoTrabajo[4] == '-'
            ? string.Concat(periodoTrabajo.AsSpan(0, 4), periodoTrabajo.AsSpan(5, 2))
            : periodoTrabajo;
    }
}
