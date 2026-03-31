using Microsoft.AspNetCore.Mvc.Rendering;
using System.Text.RegularExpressions;

namespace SistemaControlEspaciosDeportivosWeb.Services;

public static class TelefonoInternacionalHelper
{
    private static readonly (string Codigo, string Pais)[] CodigosPaisBase =
    [
        ("+51", "Peru"),
        ("+52", "Mexico"),
        ("+54", "Argentina"),
        ("+55", "Brasil"),
        ("+56", "Chile"),
        ("+57", "Colombia"),
        ("+58", "Venezuela"),
        ("+591", "Bolivia"),
        ("+593", "Ecuador"),
        ("+595", "Paraguay"),
        ("+598", "Uruguay"),
        ("+1", "EE.UU/Canada"),
        ("+34", "Espana")
    ];

    public static List<SelectListItem> ObtenerCodigosPais(string? seleccionado)
    {
        return CodigosPaisBase
            .Select(x => new SelectListItem($"{x.Pais} ({x.Codigo})", x.Codigo))
            .ToList();
    }

    public static void Descomponer(string? telefonoCompleto, out string codigoPais, out string numeroLocal)
    {
        codigoPais = "+51";
        numeroLocal = string.Empty;

        if (string.IsNullOrWhiteSpace(telefonoCompleto))
            return;

        var texto = telefonoCompleto.Trim();
        var tieneMas = texto.StartsWith('+');
        var soloDigitos = Regex.Replace(texto, @"\D", string.Empty);

        if (string.IsNullOrWhiteSpace(soloDigitos))
            return;

        if (tieneMas)
        {
            var match = CodigosPaisBase
                .OrderByDescending(x => x.Codigo.Length)
                .FirstOrDefault(x => texto.StartsWith(x.Codigo, StringComparison.Ordinal));

            if (!string.IsNullOrWhiteSpace(match.Codigo))
            {
                codigoPais = match.Codigo;
                var digitosCodigo = match.Codigo.TrimStart('+');
                numeroLocal = soloDigitos.StartsWith(digitosCodigo, StringComparison.Ordinal)
                    ? soloDigitos[digitosCodigo.Length..]
                    : soloDigitos;
                return;
            }
        }

        numeroLocal = soloDigitos;
    }

    public static string? Componer(string? codigoPais, string? numeroLocal)
    {
        var numero = Regex.Replace(numeroLocal ?? string.Empty, @"\D", string.Empty);
        if (string.IsNullOrWhiteSpace(numero))
            return null;

        var codigo = NormalizarCodigo(codigoPais);
        return $"{codigo}{numero}";
    }

    private static string NormalizarCodigo(string? codigoPais)
    {
        var limpio = Regex.Replace(codigoPais ?? string.Empty, @"[^\d+]", string.Empty);
        if (string.IsNullOrWhiteSpace(limpio))
            return "+51";
        if (!limpio.StartsWith('+'))
            limpio = "+" + limpio;
        return limpio;
    }
}
