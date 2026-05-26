USE [DbSportCenter]
GO
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
-- =============================================
-- Author:        FRANCO LARA
-- Create date:   26/05/2026
-- Description:   Inserta o actualiza avance del onboarding por negocio.
-- =============================================
-- Firma: Codex - 26/05/2026 | Se crea SP para registrar avance del wizard sin marcar finalizacion completa.
CREATE OR ALTER PROCEDURE [dbo].[Sp_OnboardingEstado_GuardarAvance]
    @NegocioId INT,
    @PasoActual TINYINT,
    @Usuario NVARCHAR(200) = NULL
AS
BEGIN
    SET NOCOUNT ON;

    BEGIN TRY
        IF @PasoActual < 1 SET @PasoActual = 1;
        IF @PasoActual > 5 SET @PasoActual = 5;

        IF EXISTS (SELECT 1 FROM dbo.NegocioOnboardingEstado WHERE NegocioId = @NegocioId)
        BEGIN
            UPDATE dbo.NegocioOnboardingEstado
            SET
                PasoActual = @PasoActual,
                FechaUltimaActualizacionUtc = SYSUTCDATETIME(),
                UsuarioUltimaActualizacion = @Usuario,
                Completado = CASE WHEN Completado = 1 THEN 1 ELSE 0 END
            WHERE NegocioId = @NegocioId;
        END
        ELSE
        BEGIN
            INSERT INTO dbo.NegocioOnboardingEstado
            (
                NegocioId,
                PasoActual,
                Completado,
                FechaUltimaActualizacionUtc,
                UsuarioUltimaActualizacion,
                FechaCompletadoUtc,
                UsuarioCompletado
            )
            VALUES
            (
                @NegocioId,
                @PasoActual,
                0,
                SYSUTCDATETIME(),
                @Usuario,
                NULL,
                NULL
            );
        END;
    END TRY
    BEGIN CATCH
        DECLARE @ErrorMessage NVARCHAR(4000), @ErrorSeverity INT, @ErrorState INT;
        SELECT @ErrorMessage = ERROR_MESSAGE(), @ErrorSeverity = ERROR_SEVERITY(), @ErrorState = ERROR_STATE();
        RAISERROR (@ErrorMessage, @ErrorSeverity, @ErrorState);
    END CATCH
END
GO
