using System.Globalization;
using System.Text;
using System.Net;
using Microsoft.Playwright;
using SistemaControlEspaciosDeportivosWeb.ViewModels;

namespace SistemaControlEspaciosDeportivosWeb.Services;

public static class InternalReceiptHtmlPdfBuilder
{
    public static async Task<byte[]> BuildAsync(ComprobanteVisualizacionViewModel model)
    {
        var html = BuildHtml(model);

        try
        {
            using var playwright = await Playwright.CreateAsync();
            await using var browser = await playwright.Chromium.LaunchAsync(new BrowserTypeLaunchOptions
            {
                Headless = true,
                Args = new[] { "--no-sandbox" }
            });

            var page = await browser.NewPageAsync(new BrowserNewPageOptions
            {
                ViewportSize = new() { Width = 1240, Height = 1754 }
            });

            await page.SetContentAsync(html, new PageSetContentOptions
            {
                WaitUntil = WaitUntilState.NetworkIdle
            });

            return await page.PdfAsync(new PagePdfOptions
            {
                Format = "A4",
                PrintBackground = true,
                Margin = new Margin
                {
                    Top = "16mm",
                    Right = "12mm",
                    Bottom = "16mm",
                    Left = "12mm"
                }
            });
        }
        catch
        {
            // Fallback para no bloquear la operacion si el runtime de Playwright no esta instalado.
            return InternalReceiptPdfBuilder.Build(model);
        }
    }

    private static string BuildHtml(ComprobanteVisualizacionViewModel model)
    {
        var numeroComprobante = $"{model.Serie}-{model.Numero:D8}";
        var ubigeoNegocio = $"{Safe(model.NegocioDepartamento)} / {Safe(model.NegocioProvincia)} / {Safe(model.NegocioDistrito)}";
        var ubigeoCliente = $"{Safe(model.ClienteDepartamento)} / {Safe(model.ClienteProvincia)} / {Safe(model.ClienteDistrito)}";
        var detalle = $"Reserva de {Safe(model.EspacioNombre)} | {model.FechaReserva:dd/MM/yyyy} {model.HoraInicioReserva:HH\\:mm}-{model.HoraFinReserva:HH\\:mm}";

        static string Monto(string simbolo, decimal valor) => $"{simbolo} {valor.ToString("N2", CultureInfo.InvariantCulture)}";

        var sb = new StringBuilder();
        sb.AppendLine("<!doctype html>");
        sb.AppendLine("<html lang=\"es\"><head><meta charset=\"utf-8\" />");
        sb.AppendLine($"<title>Recibo interno {Html(numeroComprobante)}</title>");
        sb.AppendLine("""
<style>
body { font-family: 'Segoe UI', Tahoma, Arial, sans-serif; color: #16345f; margin: 0; background: #ffffff; font-size: 12px; }
.sheet { width: 100%; max-width: 980px; margin: 0 auto; }
.header { border: 1px solid #d6e4f8; border-radius: 12px; padding: 14px; background: linear-gradient(180deg,#f3f8ff 0%,#eaf3ff 100%); }
.header h1 { margin: 0; font-size: 22px; font-weight: 800; }
.subtitle { margin-top: 4px; color: #4b658a; font-weight: 600; }
.grid { display: grid; grid-template-columns: repeat(4, 1fr); gap: 8px; margin-top: 10px; }
.card { border: 1px solid #d6e4f8; border-radius: 10px; padding: 8px 10px; background: #f8fbff; }
.card small { display: block; text-transform: uppercase; font-size: 10px; color: #6a7e9f; font-weight: 700; letter-spacing: .04em; }
.card strong { display: block; margin-top: 2px; font-size: 15px; color: #183760; }
.section { margin-top: 10px; border: 1px solid #d6e4f8; border-radius: 12px; overflow: hidden; }
.section-title { background: #eef4ff; border-bottom: 1px solid #d6e4f8; padding: 8px 10px; font-weight: 800; color: #21406a; }
.section-body { padding: 10px; }
.split { display: grid; grid-template-columns: 1fr 1fr; gap: 10px; }
.list-row { display: grid; grid-template-columns: 130px 1fr; gap: 6px; margin: 3px 0; }
.list-row .k { color: #647a9d; font-size: 11px; text-transform: uppercase; font-weight: 700; }
.list-row .v { font-weight: 700; color: #183760; word-break: break-word; }
.detalle { border: 1px solid #dce7f7; border-radius: 10px; padding: 10px; background: #fff; }
.detalle-head { display: flex; justify-content: space-between; color: #6a7e9f; font-weight: 700; font-size: 11px; margin-bottom: 6px; }
.detalle-desc { font-weight: 700; margin-bottom: 8px; color: #173a67; }
.detalle-foot { display: grid; grid-template-columns: 1fr 1fr; gap: 10px; }
.detalle-foot .m { text-align: right; }
.detalle-foot small { display: block; text-transform: uppercase; color: #647a9d; font-size: 10px; font-weight: 700; }
.detalle-foot strong { font-size: 14px; }
.totales { display: grid; grid-template-columns: 1fr; justify-items: end; }
.total-box { min-width: 260px; border: 1px solid #8eb8ff; border-radius: 10px; background: #e9f2ff; padding: 8px 10px; text-align: right; }
.total-box small { display: block; text-transform: uppercase; color: #5479ad; font-size: 10px; font-weight: 800; }
.total-box strong { font-size: 18px; color: #16345f; }
</style></head><body><main class="sheet">
""");

        sb.AppendLine("<section class=\"header\">");
        sb.AppendLine("<h1>Recibo interno</h1>");
        sb.AppendLine($"<div class=\"subtitle\">{Html(model.TipoDocumentoNombre)} | {Html(numeroComprobante)}</div>");
        sb.AppendLine("<div class=\"grid\">");
        sb.AppendLine($"<div class=\"card\"><small>Documento</small><strong>{Html(numeroComprobante)}</strong></div>");
        sb.AppendLine($"<div class=\"card\"><small>Fecha emision</small><strong>{model.FechaEmision:dd/MM/yyyy}</strong></div>");
        sb.AppendLine($"<div class=\"card\"><small>Reserva</small><strong>#{model.ReservaId}</strong></div>");
        sb.AppendLine($"<div class=\"card\"><small>Total</small><strong>{Html(Monto(model.MonedaSimbolo, model.Total))}</strong></div>");
        sb.AppendLine("</div></section>");

        sb.AppendLine("<section class=\"section\"><div class=\"section-title\">Datos del emisor y cliente</div><div class=\"section-body split\">");
        sb.AppendLine("<div>");
        sb.AppendLine($"<div class=\"list-row\"><div class=\"k\">Negocio</div><div class=\"v\">{Html(model.NegocioNombre)}</div></div>");
        sb.AppendLine("<div class=\"list-row\"><div class=\"k\">Documento</div><div class=\"v\">-</div></div>");
        sb.AppendLine($"<div class=\"list-row\"><div class=\"k\">Direccion</div><div class=\"v\">{Html(Safe(model.NegocioDireccionFiscal))}</div></div>");
        sb.AppendLine($"<div class=\"list-row\"><div class=\"k\">Ubigeo</div><div class=\"v\">{Html(ubigeoNegocio)}</div></div>");
        sb.AppendLine("</div><div>");
        sb.AppendLine($"<div class=\"list-row\"><div class=\"k\">Cliente</div><div class=\"v\">{Html(model.ClienteNombre)}</div></div>");
        sb.AppendLine($"<div class=\"list-row\"><div class=\"k\">Documento</div><div class=\"v\">{Html(Safe(model.ClienteDocumento))}</div></div>");
        sb.AppendLine($"<div class=\"list-row\"><div class=\"k\">Correo</div><div class=\"v\">{Html(Safe(model.ClienteCorreo))}</div></div>");
        sb.AppendLine($"<div class=\"list-row\"><div class=\"k\">Direccion</div><div class=\"v\">{Html(Safe(model.ClienteDireccion))}</div></div>");
        sb.AppendLine($"<div class=\"list-row\"><div class=\"k\">Ubigeo</div><div class=\"v\">{Html(ubigeoCliente)}</div></div>");
        sb.AppendLine("</div></div></section>");

        sb.AppendLine("<section class=\"section\"><div class=\"section-title\">Detalle del comprobante</div><div class=\"section-body\"><div class=\"detalle\">");
        sb.AppendLine("<div class=\"detalle-head\"><span>Item 1</span><span>UND x 1</span></div>");
        sb.AppendLine($"<div class=\"detalle-desc\">{Html(detalle)}</div>");
        sb.AppendLine("<div class=\"detalle-foot\">");
        sb.AppendLine($"<div class=\"m\"><small>Valor unitario</small><strong>{Html(Monto(model.MonedaSimbolo, model.Total))}</strong></div>");
        sb.AppendLine($"<div class=\"m\"><small>Importe</small><strong>{Html(Monto(model.MonedaSimbolo, model.Total))}</strong></div>");
        sb.AppendLine("</div></div></div></section>");

        sb.AppendLine("<section class=\"section\"><div class=\"section-title\">Resumen de importes</div><div class=\"section-body totales\">");
        sb.AppendLine($"<div class=\"total-box\"><small>Total</small><strong>{Html(Monto(model.MonedaSimbolo, model.Total))}</strong></div>");
        sb.AppendLine("</div></section></main></body></html>");
        return sb.ToString();
    }

    private static string Safe(string? value) => string.IsNullOrWhiteSpace(value) ? "-" : value.Trim();
    private static string Html(string? value) => WebUtility.HtmlEncode(value ?? string.Empty);
}
