-- =============================================
-- Author:        FRANCO LARA
-- Create date:   19/04/2026
-- Firma:         Finalizacion manual del contrato de suscripcion para dejar al negocio sin plan activo.
-- =============================================
CREATE OR ALTER PROCEDURE dbo.Sp_NegociosSuscripcion_FinalizarPlan
    @NegocioId INT,
    @Usuario NVARCHAR(200) = NULL
AS
BEGIN
    SET NOCOUNT ON;

    BEGIN TRY
        IF NOT EXISTS (SELECT 1 FROM dbo.Negocios WHERE Id = @NegocioId)
            RAISERROR('Negocio no encontrado.', 16, 1);

        UPDATE dbo.NegociosSuscripcion
        SET EstadoSuscripcion = 4,
            EsPrueba = 0,
            FechaInicioPrueba = NULL,
            FechaFinPrueba = NULL,
            FechaInicioPlan = NULL,
            FechaFinPlan = NULL,
            TipoCobro = NULL,
            FechaFinGracia = NULL,
            FechaActualizacion = SYSUTCDATETIME(),
            UsuarioActualizacion = COALESCE(@Usuario, N'sistema')
        WHERE NegocioId = @NegocioId;

        IF @@ROWCOUNT = 0
            RAISERROR('No se encontro una suscripcion para finalizar.', 16, 1);
    END TRY

    BEGIN CATCH
        DECLARE @ErrorMessage NVARCHAR(4000);
        DECLARE @ErrorSeverity INT;
        DECLARE @ErrorState INT;

        SELECT
            @ErrorMessage = ERROR_MESSAGE(),
            @ErrorSeverity = ERROR_SEVERITY(),
            @ErrorState = ERROR_STATE();

        RAISERROR (@ErrorMessage, @ErrorSeverity, @ErrorState);
    END CATCH
END
