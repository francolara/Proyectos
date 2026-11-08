// Firma: FRANCO LARA - 06/11/2026 | Firma JPG/PNG privada en R2; límites y metadatos de integridad.
using System.Buffers.Binary;
using System.Text.Json;
using GestionServicios.Application.Followup;
using GestionServicios.Application.Maintenance;
namespace GestionServicios.Web.Services;
public static class BudgetSignature
{
    public static DocumentAttachment? Metadata(RecordSnapshot? company)
    {
        var json=company?.Text("FirmaPresupuestoJson");
        return string.IsNullOrWhiteSpace(json)?null:JsonSerializer.Deserialize<DocumentAttachment>(json);
    }
    public static async Task ValidateAsync(IFormFile upload,CancellationToken ct)
    {
        if(upload.Length is <=0 or >2097152)throw new MaintenanceException("La firma debe ser JPG o PNG de hasta 2 MB.");
        var ext=Path.GetExtension(upload.FileName).ToLowerInvariant();
        if(ext is not (".png" or ".jpg" or ".jpeg"))throw new MaintenanceException("Selecciona una firma JPG o PNG.");
        var data=await DocumentUploadRules.ValidateAsync(upload,new("firma","FIRMA","Firma","CONFIGURACION"),Guid.NewGuid(),ct);
        var bytes=data.Bytes;int w=0,h=0;
        if(data.Metadata.TipoContenido=="image/png"&&bytes.Length>=24&&bytes.AsSpan(12,4).SequenceEqual("IHDR"u8))
        {w=BinaryPrimitives.ReadInt32BigEndian(bytes.AsSpan(16,4));h=BinaryPrimitives.ReadInt32BigEndian(bytes.AsSpan(20,4));}
        else if(data.Metadata.TipoContenido=="image/jpeg")
        {
            for(int i=2;i+8<bytes.Length;)
            {
                if(bytes[i++]!=255)break;while(i<bytes.Length&&bytes[i]==255)i++;if(i>=bytes.Length)break;
                var marker=bytes[i++];if(marker is 217 or 218)break;if(marker is 216 or 1 or >=208 and <=215)continue;
                if(i+2>bytes.Length)break;int size=BinaryPrimitives.ReadUInt16BigEndian(bytes.AsSpan(i,2));if(size<2||i+size>bytes.Length)break;
                if(marker is 192 or 193 or 194 or 195 or 197 or 198 or 199 or 201 or 202 or 203 or 205 or 206 or 207){h=BinaryPrimitives.ReadUInt16BigEndian(bytes.AsSpan(i+3,2));w=BinaryPrimitives.ReadUInt16BigEndian(bytes.AsSpan(i+5,2));break;}
                i+=size;
            }
        }
        if(w is <1 or >4096||h is <1 or >4096)throw new MaintenanceException("La firma no es una imagen válida o supera 4096 píxeles por lado.");
    }
}
