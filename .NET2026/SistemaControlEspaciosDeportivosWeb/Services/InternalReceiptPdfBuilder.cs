using System.Globalization;
using System.Text;
using SistemaControlEspaciosDeportivosWeb.ViewModels;

namespace SistemaControlEspaciosDeportivosWeb.Services;

public static class InternalReceiptPdfBuilder
{
    public static byte[] Build(ComprobanteVisualizacionViewModel model)
    {
        var lines = new List<string>
        {
            "RECIBO INTERNO",
            $"{model.Serie}-{model.Numero:D8}",
            $"Fecha emision: {model.FechaEmision:dd/MM/yyyy}",
            "",
            $"Negocio: {model.NegocioNombre}",
            $"Razon social: {Safe(model.NegocioRazonSocial)}",
            $"Documento: {Safe(model.NegocioDocumento)}",
            $"Direccion: {Safe(model.NegocioDireccionFiscal)}",
            "",
            $"Cliente: {model.ClienteNombre}",
            $"Documento cliente: {Safe(model.ClienteDocumento)}",
            $"Correo: {Safe(model.ClienteCorreo)}",
            $"Direccion cliente: {Safe(model.ClienteDireccion)}",
            "",
            $"Reserva: #{model.ReservaId} | Sede: {model.SedeNombre} | Espacio: {model.EspacioNombre}",
            $"Fecha reserva: {model.FechaReserva:dd/MM/yyyy} {model.HoraInicioReserva:HH\\:mm}-{model.HoraFinReserva:HH\\:mm}",
            "",
            "DETALLE",
            "Item: 1",
            $"Descripcion: Reserva {model.EspacioNombre} ({model.FechaReserva:dd/MM/yyyy} {model.HoraInicioReserva:HH\\:mm}-{model.HoraFinReserva:HH\\:mm})",
            "Unidad: UND",
            "Cantidad: 1",
            $"Valor unitario: {model.MonedaSimbolo} {model.Total.ToString("N2", CultureInfo.InvariantCulture)}",
            $"Importe: {model.MonedaSimbolo} {model.Total.ToString("N2", CultureInfo.InvariantCulture)}",
            "",
            $"TOTAL: {model.MonedaSimbolo} {model.Total.ToString("N2", CultureInfo.InvariantCulture)}"
        };

        return BuildSimplePdf(lines);
    }

    private static string Safe(string? value) => string.IsNullOrWhiteSpace(value) ? "-" : value.Trim();

    private static byte[] BuildSimplePdf(List<string> lines)
    {
        var objects = new List<string>();
        var xref = new List<int> { 0 };

        string Escape(string input) => input.Replace("\\", "\\\\").Replace("(", "\\(").Replace(")", "\\)");

        var content = new StringBuilder();
        content.AppendLine("BT");
        content.AppendLine("/F1 11 Tf");
        content.AppendLine("40 800 Td");
        content.AppendLine("14 TL");
        for (var i = 0; i < lines.Count; i++)
        {
            var line = Escape(lines[i]);
            if (i == 0)
            {
                content.AppendLine($"({line}) Tj");
            }
            else
            {
                content.AppendLine("T*");
                content.AppendLine($"({line}) Tj");
            }
        }
        content.AppendLine("ET");

        var stream = content.ToString();
        objects.Add("1 0 obj << /Type /Catalog /Pages 2 0 R >> endobj");
        objects.Add("2 0 obj << /Type /Pages /Count 1 /Kids [3 0 R] >> endobj");
        objects.Add("3 0 obj << /Type /Page /Parent 2 0 R /MediaBox [0 0 595 842] /Resources << /Font << /F1 4 0 R >> >> /Contents 5 0 R >> endobj");
        objects.Add("4 0 obj << /Type /Font /Subtype /Type1 /BaseFont /Helvetica >> endobj");
        objects.Add($"5 0 obj << /Length {Encoding.ASCII.GetByteCount(stream)} >> stream\n{stream}endstream endobj");

        var pdf = new StringBuilder();
        pdf.AppendLine("%PDF-1.4");
        foreach (var obj in objects)
        {
            xref.Add(Encoding.ASCII.GetByteCount(pdf.ToString()));
            pdf.AppendLine(obj);
        }

        var xrefStart = Encoding.ASCII.GetByteCount(pdf.ToString());
        pdf.AppendLine($"xref\n0 {objects.Count + 1}");
        pdf.AppendLine("0000000000 65535 f ");
        for (var i = 1; i <= objects.Count; i++)
        {
            pdf.AppendLine($"{xref[i]:D10} 00000 n ");
        }

        pdf.AppendLine($"trailer << /Size {objects.Count + 1} /Root 1 0 R >>");
        pdf.AppendLine("startxref");
        pdf.AppendLine(xrefStart.ToString(CultureInfo.InvariantCulture));
        pdf.Append("%%EOF");

        return Encoding.ASCII.GetBytes(pdf.ToString());
    }
}
