USE [DbSportCenter]
GO
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
-- =============================================
-- Author:        FRANCO LARA
-- Create date:   26/05/2026
-- Description:   Obtiene el estado de avance del onboarding por negocio.
-- =============================================
-- Firma: Codex - 26/05/2026 | Se crea SP para consultar estado de onboarding y retornar defaults cuando no exista registro.
CREATE OR ALTER PROCEDURE [dbo].[Sp_OnboardingEstado_Obtener]
    @NegocioId INT
AS
BEGIN
    SET NOCOUNT ON;

    BEGIN TRY
        SELECT
            @NegocioId AS NegocioId,
            CAST(COALESCE(o.PasoActual, 1) AS TINYINT) AS PasoActual,
            CAST(COALESCE(o.Completado, 0) AS BIT) AS Completado,
            o.FechaUltimaActualizacionUtc,
            o.UsuarioUltimaActualizacion,
            o.FechaCompletadoUtc,
            o.UsuarioCompletado
        FROM (VALUES (1)) AS v(Id)
        LEFT JOIN dbo.NegocioOnboardingEstado o
            ON o.NegocioId = @NegocioId;
    END TRY
    BEGIN CATCH
        DECLARE @ErrorMessage NVARCHAR(4000), @ErrorSeverity INT, @ErrorState INT;
        SELECT @ErrorMessage = ERROR_MESSAGE(), @ErrorSeverity = ERROR_SEVERITY(), @ErrorState = ERROR_STATE();
        RAISERROR (@ErrorMessage, @ErrorSeverity, @ErrorState);
    END CATCH
END
GO
