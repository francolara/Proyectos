USE [DbSportCenter]
GO
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
-- =============================================
-- Author:        FRANCO LARA
-- Create date:   26/05/2026
-- Description:   Marca onboarding como completado por negocio.
-- =============================================
-- Firma: Codex - 26/05/2026 | Se crea SP para finalizar onboarding y dejar paso actual en resumen final.
CREATE OR ALTER PROCEDURE [dbo].[Sp_OnboardingEstado_MarcarCompletado]
    @NegocioId INT,
    @Usuario NVARCHAR(200) = NULL
AS
BEGIN
    SET NOCOUNT ON;

    BEGIN TRY
        IF EXISTS (SELECT 1 FROM dbo.NegocioOnboardingEstado WHERE NegocioId = @NegocioId)
        BEGIN
            UPDATE dbo.NegocioOnboardingEstado
            SET
                PasoActual = 5,
                Completado = 1,
                FechaUltimaActualizacionUtc = SYSUTCDATETIME(),
                UsuarioUltimaActualizacion = @Usuario,
                FechaCompletadoUtc = SYSUTCDATETIME(),
                UsuarioCompletado = @Usuario
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
                5,
                1,
                SYSUTCDATETIME(),
                @Usuario,
                SYSUTCDATETIME(),
                @Usuario
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
