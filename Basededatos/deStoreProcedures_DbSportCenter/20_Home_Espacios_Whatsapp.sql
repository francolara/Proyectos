-- =============================================
-- Author:        FRANCO LARA
-- Create date:   27/03/2026
-- Description:   Sprint 7.3 - WhatsApp por sede en resultados de espacios disponibles.
-- =============================================

CREATE OR ALTER PROCEDURE dbo.Sp_Home_BuscarEspaciosDisponibles
    @Fecha DATE,
    @HoraInicio TIME,
    @HoraFin TIME,
    @SedeId INT = NULL,
    @TipoDeporteId INT = NULL
AS
BEGIN
    SET NOCOUNT ON;
    BEGIN TRY
        SELECT
            e.Id,
            e.Nombre,
            e.Codigo,
            s.Nombre AS SedeNombre,
            td.Nombre AS TipoDeporte,
            e.TieneIluminacion,
            e.Techada,
            scn.WhatsappContacto,
            COALESCE(scn.PermiteChatWhatsapp, 0) AS PermiteChatWhatsapp
        FROM dbo.EspaciosDeportivos e
        INNER JOIN dbo.Sedes s ON s.Id = e.SedeId
        INNER JOIN dbo.TiposDeporte td ON td.Id = e.TipoDeporteId
        LEFT JOIN dbo.SedeConfiguracionNotificacion scn ON scn.SedeId = s.Id
        WHERE e.Estado = 1
          AND s.Activo = 1
          AND (@SedeId IS NULL OR e.SedeId = @SedeId)
          AND (@TipoDeporteId IS NULL OR e.TipoDeporteId = @TipoDeporteId)
          AND NOT EXISTS
          (
              SELECT 1
              FROM dbo.Reservas r
              WHERE r.EspacioDeportivoId = e.Id
                AND r.Fecha = @Fecha
                AND r.Estado NOT IN (5, 6)
                AND @HoraInicio < r.HoraFin
                AND @HoraFin > r.HoraInicio
          )
        ORDER BY s.Nombre, e.Nombre;
    END TRY
    BEGIN CATCH
        DECLARE @ErrorMessage NVARCHAR(4000), @ErrorSeverity INT, @ErrorState INT;
        SELECT @ErrorMessage = ERROR_MESSAGE(), @ErrorSeverity = ERROR_SEVERITY(), @ErrorState = ERROR_STATE();
        RAISERROR (@ErrorMessage, @ErrorSeverity, @ErrorState);
    END CATCH
END
GO
