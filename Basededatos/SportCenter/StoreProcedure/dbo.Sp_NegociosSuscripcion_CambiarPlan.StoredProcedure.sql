-- =============================================
-- Author:        FRANCO LARA
-- Create date:   10/06/2026
-- Firma:         Cambia el tipo de plan y vigencia del contrato comercial del negocio dejando trazabilidad.
-- =============================================
CREATE OR ALTER PROCEDURE dbo.Sp_NegociosSuscripcion_CambiarPlan
    @NegocioId INT,
    @TipoCobro NVARCHAR(20),
    @FechaInicioPlan DATE,
    @FechaFinPlan DATE,
    @DiasGracia INT = 5,
    @Observacion NVARCHAR(500) = NULL,
    @Usuario NVARCHAR(200) = NULL
AS
BEGIN
    SET NOCOUNT ON;

    BEGIN TRY
        DECLARE @SuscripcionId INT;
        DECLARE @EstadoAnterior INT;
        DECLARE @EsPruebaAnterior BIT;
        DECLARE @TipoCobroAnterior NVARCHAR(20);
        DECLARE @TipoCobroNorm NVARCHAR(20) = UPPER(LTRIM(RTRIM(COALESCE(@TipoCobro, N''))));

        IF @NegocioId IS NULL OR @NegocioId <= 0
            RAISERROR('Negocio invalido.', 16, 1);

        IF @TipoCobroNorm NOT IN (N'MENSUAL', N'TRIMESTRAL', N'SEMESTRAL', N'ANUAL')
            RAISERROR('Tipo de cobro invalido. Usa MENSUAL, TRIMESTRAL, SEMESTRAL o ANUAL.', 16, 1);

        IF @FechaInicioPlan IS NULL OR @FechaFinPlan IS NULL OR @FechaFinPlan < @FechaInicioPlan
            RAISERROR('Rango de vigencia invalido para el plan.', 16, 1);

        IF @DiasGracia IS NULL OR @DiasGracia < 0
            SET @DiasGracia = 5;

        IF NOT EXISTS (SELECT 1 FROM dbo.Negocios WHERE Id = @NegocioId)
            RAISERROR('Negocio no encontrado.', 16, 1);

        SELECT
            @SuscripcionId = ns.Id,
            @EstadoAnterior = ns.EstadoSuscripcion,
            @EsPruebaAnterior = ns.EsPrueba,
            @TipoCobroAnterior = UPPER(LTRIM(RTRIM(COALESCE(ns.TipoCobro, N''))))
        FROM dbo.NegociosSuscripcion ns
        WHERE ns.NegocioId = @NegocioId;

        IF @SuscripcionId IS NULL
            RAISERROR('El negocio no tiene una suscripcion creada.', 16, 1);

        IF COALESCE(@EsPruebaAnterior, 0) = 1
            RAISERROR('No se puede cambiar el plan mientras el negocio este en prueba.', 16, 1);

        UPDATE dbo.NegociosSuscripcion
        SET EstadoSuscripcion = 2,
            EsPrueba = 0,
            FechaInicioPrueba = NULL,
            FechaFinPrueba = NULL,
            FechaInicioPlan = @FechaInicioPlan,
            FechaFinPlan = @FechaFinPlan,
            TipoCobro = @TipoCobroNorm,
            DiasGracia = @DiasGracia,
            FechaFinGracia = DATEADD(DAY, @DiasGracia, @FechaFinPlan),
            FechaActualizacion = SYSUTCDATETIME(),
            UsuarioActualizacion = COALESCE(@Usuario, N'sistema')
        WHERE NegocioId = @NegocioId;

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
            @NegocioId, @SuscripcionId, N'CAMBIO_PLAN',
            @EstadoAnterior, 2,
            @EsPruebaAnterior, 0,
            NULLIF(@TipoCobroAnterior, N''), @TipoCobroNorm,
            @FechaInicioPlan, @FechaFinPlan,
            @DiasGracia, NULL, NULLIF(LTRIM(RTRIM(COALESCE(@Observacion, N''))), N''),
            SYSUTCDATETIME(), COALESCE(@Usuario, N'sistema')
        );
    END TRY
    BEGIN CATCH
        DECLARE @ErrorMessage NVARCHAR(4000), @ErrorSeverity INT, @ErrorState INT;
        SELECT
            @ErrorMessage = ERROR_MESSAGE(),
            @ErrorSeverity = ERROR_SEVERITY(),
            @ErrorState = ERROR_STATE();
        RAISERROR (@ErrorMessage, @ErrorSeverity, @ErrorState);
    END CATCH
END
GO
