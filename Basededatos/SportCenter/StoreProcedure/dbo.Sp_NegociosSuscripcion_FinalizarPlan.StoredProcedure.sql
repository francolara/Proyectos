-- =============================================
-- Author:        FRANCO LARA
-- Create date:   19/04/2026
-- Firma:         Finalizacion manual del contrato de suscripcion para dejar al negocio sin plan activo.
-- Firma:         10/06/2026 | Registra movimiento comercial de finalizacion o suspension manual del contrato.
-- =============================================
CREATE OR ALTER PROCEDURE dbo.Sp_NegociosSuscripcion_FinalizarPlan
    @NegocioId INT,
    @Usuario NVARCHAR(200) = NULL
AS
BEGIN
    SET NOCOUNT ON;

    BEGIN TRY
        DECLARE @SuscripcionId INT;
        DECLARE @EstadoAnterior INT;
        DECLARE @EsPruebaAnterior BIT;
        DECLARE @TipoCobroAnterior NVARCHAR(20);
        DECLARE @FechaInicioReferencia DATE;
        DECLARE @FechaFinReferencia DATE;

        IF NOT EXISTS (SELECT 1 FROM dbo.Negocios WHERE Id = @NegocioId)
            RAISERROR('Negocio no encontrado.', 16, 1);

        SELECT
            @SuscripcionId = ns.Id,
            @EstadoAnterior = ns.EstadoSuscripcion,
            @EsPruebaAnterior = ns.EsPrueba,
            @TipoCobroAnterior = UPPER(LTRIM(RTRIM(COALESCE(ns.TipoCobro, N'')))),
            @FechaInicioReferencia = COALESCE(ns.FechaInicioPlan, ns.FechaInicioPrueba),
            @FechaFinReferencia = COALESCE(ns.FechaFinPlan, ns.FechaFinPrueba)
        FROM dbo.NegociosSuscripcion ns
        WHERE ns.NegocioId = @NegocioId;

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

        IF OBJECT_ID(N'dbo.NegociosSuscripcionMovimiento', N'U') IS NOT NULL
        BEGIN
            INSERT INTO dbo.NegociosSuscripcionMovimiento
            (
                NegocioId, NegocioSuscripcionId, TipoMovimiento,
                EstadoSuscripcionAnterior, EstadoSuscripcionNuevo,
                EsPruebaAnterior, EsPruebaNuevo,
                TipoCobroAnterior, TipoCobroNuevo,
                FechaInicioReferencia, FechaFinReferencia,
                DiasGracia, DiasExtra, Observacion,
                FechaCreacion, UsuarioCreacion
            )
            VALUES
            (
                @NegocioId, @SuscripcionId, N'FINALIZACION',
                @EstadoAnterior, 4,
                @EsPruebaAnterior, 0,
                NULLIF(@TipoCobroAnterior, N''), NULL,
                @FechaInicioReferencia, @FechaFinReferencia,
                NULL, NULL, N'Finalizacion manual del contrato desde superadmin.',
                SYSUTCDATETIME(), COALESCE(@Usuario, N'sistema')
            );
        END;
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
