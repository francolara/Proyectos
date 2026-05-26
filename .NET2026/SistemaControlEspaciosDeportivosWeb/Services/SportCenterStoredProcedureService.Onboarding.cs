using System.Data;
using Microsoft.Data.SqlClient;
using SistemaControlEspaciosDeportivosWeb.ViewModels;

namespace SistemaControlEspaciosDeportivosWeb.Services;

public partial class SportCenterStoredProcedureService
{
    public async Task<OnboardingChecklistViewModel> OnboardingChecklistValidarAsync(int negocioId)
    {
        await using var cn = CreateConnection();
        await cn.OpenAsync();
        await using var cmd = new SqlCommand("Sp_OnboardingChecklist_Validar", cn) { CommandType = CommandType.StoredProcedure };
        AddParam(cmd, "@NegocioId", negocioId, SqlDbType.Int);
        await using var dr = await cmd.ExecuteReaderAsync();
        if (!await dr.ReadAsync())
            return new OnboardingChecklistViewModel { NegocioId = negocioId };

        return new OnboardingChecklistViewModel
        {
            NegocioId = dr.IsDBNull(0) ? negocioId : dr.GetInt32(0),
            ConfigNombreComercialOk = !dr.IsDBNull(1) && Convert.ToBoolean(dr.GetValue(1)),
            ConfigTipoDocumentoOk = !dr.IsDBNull(2) && Convert.ToBoolean(dr.GetValue(2)),
            ConfigMonedaOk = !dr.IsDBNull(3) && Convert.ToBoolean(dr.GetValue(3)),
            ConfigCpeCondicionesOk = !dr.IsDBNull(4) && Convert.ToBoolean(dr.GetValue(4)),
            MaestroTipoDeporteOk = !dr.IsDBNull(5) && Convert.ToBoolean(dr.GetValue(5)),
            MaestroTipoSueloOk = !dr.IsDBNull(6) && Convert.ToBoolean(dr.GetValue(6)),
            MaestroFormaPagoOk = !dr.IsDBNull(7) && Convert.ToBoolean(dr.GetValue(7)),
            MaestroMonedaOk = !dr.IsDBNull(8) && Convert.ToBoolean(dr.GetValue(8)),
            MaestroTipoDocumentoOk = !dr.IsDBNull(9) && Convert.ToBoolean(dr.GetValue(9)),
            MaestroSerieDocumentoOk = !dr.IsDBNull(10) && Convert.ToBoolean(dr.GetValue(10)),
            SedeMinimaOk = !dr.IsDBNull(11) && Convert.ToBoolean(dr.GetValue(11)),
            EspacioMinimoOk = !dr.IsDBNull(12) && Convert.ToBoolean(dr.GetValue(12)),
            ChecklistCompleto = !dr.IsDBNull(13) && Convert.ToBoolean(dr.GetValue(13))
        };
    }
}
