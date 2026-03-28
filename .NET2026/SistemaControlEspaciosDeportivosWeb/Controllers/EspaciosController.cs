using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.Mvc.Rendering;
using SistemaControlEspaciosDeportivosWeb.Services;
using SistemaControlEspaciosDeportivosWeb.ViewModels;
using System.Text.Json;

namespace SistemaControlEspaciosDeportivosWeb.Controllers;

public class EspaciosController(IModuloPermisoService moduloPermisoService, ISportCenterStoredProcedureService spService)
    : ModuloControllerBase(moduloPermisoService)
{
    public async Task<IActionResult> Index(int? negocioId)
    {
        var resolvedNegocioId = await ResolverNegocioIdAsync(negocioId, spService);
        if (!resolvedNegocioId.HasValue) return Forbid();

        var baseVm = await ObtenerBaseAsync(resolvedNegocioId.Value, "ESPACIOS");
        if (baseVm is null || !string.IsNullOrWhiteSpace(baseVm.Mensaje)) return SinAcceso(baseVm ?? new ModuloBaseViewModel { Mensaje = "Acceso denegado." });

        var vm = new EspaciosIndexViewModel
        {
            NegocioId = baseVm.NegocioId,
            NegocioNombre = baseVm.NegocioNombre,
            RolActual = baseVm.RolActual,
            ModuloCodigo = baseVm.ModuloCodigo,
            ModuloNombre = baseVm.ModuloNombre,
            PuedeCrear = baseVm.PuedeCrear,
            PuedeEditar = baseVm.PuedeEditar,
            PuedeEliminar = baseVm.PuedeEliminar,
            Espacios = await spService.EspaciosListarAsync(resolvedNegocioId.Value)
        };
        return View(vm);
    }

    public async Task<IActionResult> Create(int? negocioId)
    {
        var resolvedNegocioId = await ResolverNegocioIdAsync(negocioId, spService);
        if (!resolvedNegocioId.HasValue) return Forbid();

        var baseVm = await ObtenerBaseAsync(resolvedNegocioId.Value, "ESPACIOS");
        if (baseVm is null || !baseVm.PuedeCrear) return SinAcceso(baseVm ?? new ModuloBaseViewModel { Mensaje = "No autorizado." });

        var vm = new EspacioFormViewModel { NegocioId = resolvedNegocioId.Value, NegocioNombre = baseVm.NegocioNombre, RolActual = baseVm.RolActual };
        await CargarCombosEspacioAsync(vm);
        return View(vm);
    }

    [HttpPost]
    [ValidateAntiForgeryToken]
    public async Task<IActionResult> Create(EspacioFormViewModel model)
    {
        var baseVm = await ObtenerBaseAsync(model.NegocioId, "ESPACIOS");
        if (baseVm is null || !baseVm.PuedeCrear) return SinAcceso(baseVm ?? new ModuloBaseViewModel { Mensaje = "No autorizado." });

        await CargarCombosEspacioAsync(model);
        CargarTarifasDesdeJson(model);
        ValidarTarifas(model);
        if (!ModelState.IsValid) return View(model);

        await spService.EspaciosCrearAsync(model, User.Identity?.Name ?? "sistema");
        return RedirectToAction(nameof(Index), new { negocioId = model.NegocioId });
    }

    public async Task<IActionResult> Edit(int id, int? negocioId)
    {
        var resolvedNegocioId = await ResolverNegocioIdAsync(negocioId, spService);
        if (!resolvedNegocioId.HasValue) return Forbid();

        var baseVm = await ObtenerBaseAsync(resolvedNegocioId.Value, "ESPACIOS");
        if (baseVm is null || !baseVm.PuedeEditar) return SinAcceso(baseVm ?? new ModuloBaseViewModel { Mensaje = "No autorizado." });

        var vm = await spService.EspaciosObtenerAsync(resolvedNegocioId.Value, id);
        if (vm is null) return NotFound();
        vm.NegocioNombre = baseVm.NegocioNombre;
        vm.RolActual = baseVm.RolActual;
        await CargarCombosEspacioAsync(vm);
        return View(vm);
    }

    [HttpPost]
    [ValidateAntiForgeryToken]
    public async Task<IActionResult> Edit(EspacioFormViewModel model)
    {
        var baseVm = await ObtenerBaseAsync(model.NegocioId, "ESPACIOS");
        if (baseVm is null || !baseVm.PuedeEditar) return SinAcceso(baseVm ?? new ModuloBaseViewModel { Mensaje = "No autorizado." });

        await CargarCombosEspacioAsync(model);
        CargarTarifasDesdeJson(model);
        ValidarTarifas(model);
        if (!ModelState.IsValid) return View(model);

        var ok = await spService.EspaciosActualizarAsync(model, User.Identity?.Name ?? "sistema");
        if (!ok)
        {
            ModelState.AddModelError(string.Empty, "No se pudo guardar el espacio deportivo. Verifica el negocio seleccionado.");
            return View(model);
        }
        return RedirectToAction(nameof(Index), new { negocioId = model.NegocioId });
    }

    [HttpPost]
    [ValidateAntiForgeryToken]
    public async Task<IActionResult> Delete(int negocioId, int id)
    {
        var baseVm = await ObtenerBaseAsync(negocioId, "ESPACIOS");
        if (baseVm is null || !baseVm.PuedeEliminar) return SinAcceso(baseVm ?? new ModuloBaseViewModel { Mensaje = "No autorizado." });

        var ok = await spService.EspaciosEliminarAsync(negocioId, id, User.Identity?.Name ?? "sistema");
        if (!ok) return NotFound();
        return RedirectToAction(nameof(Index), new { negocioId });
    }

    private async Task CargarCombosEspacioAsync(EspacioFormViewModel model)
    {
        model.Sedes = await spService.EspaciosComboSedesAsync(model.NegocioId);
        model.TiposDeporte = await spService.EspaciosComboTiposDeporteAsync();
        model.TiposSuelo = await spService.EspaciosComboTiposSueloAsync();
        model.TarifaDiasSemana =
        [
            new("Lunes", "1"),
            new("Martes", "2"),
            new("Miercoles", "3"),
            new("Jueves", "4"),
            new("Viernes", "5"),
            new("Sabado", "6"),
            new("Domingo", "0")
        ];

        if (model.Tarifas.Count == 0)
        {
            model.Tarifas =
            [
                new EspacioTarifaRangoViewModel
                {
                    DiaSemana = 1,
                    HoraInicio = new TimeOnly(8, 0),
                    HoraFin = new TimeOnly(9, 0),
                    Precio = 0.01m
                }
            ];
        }

        model.TarifasJson = JsonSerializer.Serialize(model.Tarifas);
    }

    private void CargarTarifasDesdeJson(EspacioFormViewModel model)
    {
        if (string.IsNullOrWhiteSpace(model.TarifasJson))
            return;

        try
        {
            var tarifas = JsonSerializer.Deserialize<List<EspacioTarifaRangoViewModel>>(model.TarifasJson);
            if (tarifas is { Count: > 0 })
                model.Tarifas = tarifas;
        }
        catch
        {
            ModelState.AddModelError(string.Empty, "No se pudo leer el detalle de tarifas.");
        }
    }

    private void ValidarTarifas(EspacioFormViewModel model)
    {
        if (model.Tarifas.Count == 0)
        {
            ModelState.AddModelError(string.Empty, "Debes registrar al menos una tarifa.");
            return;
        }

        foreach (var tarifa in model.Tarifas)
        {
            if (tarifa.DiaSemana < 0 || tarifa.DiaSemana > 6)
                ModelState.AddModelError(string.Empty, "Existe una tarifa con dia invalido.");

            if (tarifa.HoraFin <= tarifa.HoraInicio)
                ModelState.AddModelError(string.Empty, "En tarifas, la hora fin debe ser mayor que la hora inicio.");

            if (tarifa.Precio <= 0)
                ModelState.AddModelError(string.Empty, "En tarifas, el precio debe ser mayor que cero.");
        }

        var tarifasPorDia = model.Tarifas.GroupBy(t => t.DiaSemana);
        foreach (var grupoDia in tarifasPorDia)
        {
            var ordenado = grupoDia.OrderBy(t => t.HoraInicio).ToList();
            for (var i = 1; i < ordenado.Count; i++)
            {
                if (ordenado[i].HoraInicio < ordenado[i - 1].HoraFin)
                {
                    ModelState.AddModelError(string.Empty, "Existen rangos de tarifa superpuestos en el mismo dia.");
                    break;
                }
            }
        }

        model.TarifasJson = JsonSerializer.Serialize(model.Tarifas);
    }
}
