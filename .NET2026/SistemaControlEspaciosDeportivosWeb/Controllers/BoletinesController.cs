using Microsoft.AspNetCore.Authorization;
using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.Mvc.Rendering;
using SistemaControlEspaciosDeportivosWeb.Services;
using SistemaControlEspaciosDeportivosWeb.ViewModels;

namespace SistemaControlEspaciosDeportivosWeb.Controllers;

[AllowAnonymous]
public class BoletinesController(ISportCenterStoredProcedureService spService) : Controller
{
    [HttpGet]
    public async Task<IActionResult> Index(string? codigoDepartamento = null, string? codigoProvincia = null, string? codigoUbigeo = null, string? zona = null, int? anio = null, int? mes = null)
    {
        ViewData["PublicFullWidth"] = true;

        var hoy = DateTime.Today;
        var vm = new BoletinesPublicosIndexViewModel
        {
            CodigoDepartamento = string.IsNullOrWhiteSpace(codigoDepartamento) ? null : codigoDepartamento.Trim(),
            CodigoProvincia = string.IsNullOrWhiteSpace(codigoProvincia) ? null : codigoProvincia.Trim(),
            CodigoUbigeo = string.IsNullOrWhiteSpace(codigoUbigeo) ? null : codigoUbigeo.Trim(),
            Zona = string.IsNullOrWhiteSpace(zona) ? null : zona.Trim(),
            Anio = anio ?? hoy.Year,
            Mes = mes ?? hoy.Month
        };

        vm.Departamentos = await spService.UbigeoDepartamentosListarAsync();
        vm.Provincias = !string.IsNullOrWhiteSpace(vm.CodigoDepartamento) && vm.CodigoDepartamento.Length == 2
            ? await spService.UbigeoProvinciasListarAsync(vm.CodigoDepartamento)
            : new List<SelectListItem>();
        vm.Distritos = !string.IsNullOrWhiteSpace(vm.CodigoProvincia) && vm.CodigoProvincia.Length == 4
            ? await spService.UbigeoDistritosListarAsync(vm.CodigoProvincia)
            : new List<SelectListItem>();
        vm.Zonas = await spService.UbigeoZonasListarAsync(vm.CodigoDepartamento, vm.CodigoProvincia);
        vm.Anios = Enumerable.Range(hoy.Year - 1, 4)
            .Select(x => new SelectListItem(x.ToString(), x.ToString(), x == vm.Anio))
            .ToList();
        vm.Meses = Enumerable.Range(1, 12)
            .Select(x => new SelectListItem(System.Globalization.CultureInfo.GetCultureInfo("es-PE").DateTimeFormat.GetMonthName(x), x.ToString(), x == vm.Mes))
            .ToList();
        vm.Boletines = await spService.BoletinesDeportivosListarPublicoAsync(vm.CodigoDepartamento, vm.CodigoProvincia, vm.CodigoUbigeo, vm.Zona, vm.Anio, vm.Mes);

        return View(vm);
    }

    [HttpGet]
    public async Task<IActionResult> Detalle(int id)
    {
        ViewData["PublicFullWidth"] = true;
        var boletin = await spService.BoletinesDeportivosObtenerPorIdAsync(id);
        if (boletin is null || !boletin.Activo)
            return NotFound();

        return View(boletin);
    }
}
