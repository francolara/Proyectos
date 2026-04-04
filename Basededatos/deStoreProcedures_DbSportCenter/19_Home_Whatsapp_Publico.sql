-- =============================================
-- Author:        FRANCO LARA
-- Create date:   27/03/2026
-- Description:   Sprint 7.2 - Exposicion de WhatsApp de sede en portal publico.
-- Firma:         Codex - 02/04/2026 | Expone ubicacion y fotos de sede (principal+alternativas) para portal publico.
-- =============================================

CREATE OR ALTER PROCEDURE dbo.Sp_Home_ListarSedesPublicas
AS
BEGIN
    SET NOCOUNT ON;
    BEGIN TRY
        SELECT
            s.Id,
            s.Nombre,
            s.Direccion,
            s.Telefono,
            scn.WhatsappContacto,
            COALESCE(scn.PermiteChatWhatsapp, 0) AS PermiteChatWhatsapp,
            s.Latitud,
            s.Longitud,
            s.GoogleMapsUrl,
            s.FotoPrincipalUrl,
            s.FotosUrlsCsv
        FROM dbo.Sedes s
        INNER JOIN dbo.Negocios n ON n.Id = s.NegocioId
        LEFT JOIN dbo.SedeConfiguracionNotificacion scn ON scn.SedeId = s.Id
        WHERE s.Activo = 1
          AND n.Activo = 1
        ORDER BY s.Nombre;
    END TRY
    BEGIN CATCH
        DECLARE @ErrorMessage NVARCHAR(4000), @ErrorSeverity INT, @ErrorState INT;
        SELECT @ErrorMessage = ERROR_MESSAGE(), @ErrorSeverity = ERROR_SEVERITY(), @ErrorState = ERROR_STATE();
        RAISERROR (@ErrorMessage, @ErrorSeverity, @ErrorState);
    END CATCH
END
GO
