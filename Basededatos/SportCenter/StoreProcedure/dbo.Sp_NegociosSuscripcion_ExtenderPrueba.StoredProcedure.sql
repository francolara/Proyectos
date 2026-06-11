-- =============================================
-- Author:        FRANCO LARA
-- Create date:   10/06/2026
-- Firma:         Extiende o reactiva periodo de prueba del negocio y registra movimiento comercial.
-- =============================================
CREATE OR ALTER PROCEDURE dbo.Sp_NegociosSuscripcion_ExtenderPrueba
    @NegocioId INT,
    @DiasExtra INT,
    @Observacion NVARCHAR(500) = NULL,
    @Usuario NVARCHAR(200) = NULL
AS
BEGIN
    SET NOCOUNT ON;

    BEGIN TRY
        DECLARE @SuscripcionId INT;
        DECLARE @EstadoAnterior INT;
        DECLARE @EsPruebaAnterior BIT;
        DECLARE @FechaInicioPrueba DATE;
        DECLARE @FechaFinPruebaAnterior DATE;
        DECLARE @NuevaFechaFinPrueba DATE;
        DECLARE @Hoy DATE = CAST(GETDATE() AS DATE);

        IF @NegocioId IS NULL OR @NegocioId <= 0
            RAISERROR('Negocio invalido.', 16, 1);

        IF @DiasExtra IS NULL OR @DiasExtra <= 0
            RAISERROR('Ingresa la cantidad de dias extra para la prueba.', 16, 1);

        IF NOT EXISTS (SELECT 1 FROM dbo.Negocios WHERE Id = @NegocioId)
            RAISERROR('Negocio no encontrado.', 16, 1);

        SELECT
            @SuscripcionId = ns.Id,
            @EstadoAnterior = ns.EstadoSuscripcion,
            @EsPruebaAnterior = ns.EsPrueba,
            @FechaInicioPrueba = ns.FechaInicioPrueba,
            @FechaFinPruebaAnterior = ns.FechaFinPrueba
        FROM dbo.NegociosSuscripcion ns
        WHERE ns.NegocioId = @NegocioId;

        IF @SuscripcionId IS NULL
            RAISERROR('El negocio no tiene una suscripcion creada para extender la prueba.', 16, 1);

        IF COALESCE(@EsPruebaAnterior, 0) = 0 AND @FechaFinPruebaAnterior IS NULL
            RAISERROR('El negocio no tiene una prueba previa para extender.', 16, 1);

        IF @FechaInicioPrueba IS NULL
            SET @FechaInicioPrueba = @Hoy;

        SET @NuevaFechaFinPrueba = DATEADD(DAY, @DiasExtra, CASE
            WHEN @FechaFinPruebaAnterior IS NULL THEN @Hoy
            WHEN @FechaFinPruebaAnterior < @Hoy THEN @Hoy
            ELSE @FechaFinPruebaAnterior
        END);

        UPDATE dbo.NegociosSuscripcion
        SET EstadoSuscripcion = 1,
            EsPrueba = 1,
            FechaInicioPrueba = @FechaInicioPrueba,
            FechaFinPrueba = @NuevaFechaFinPrueba,
            FechaInicioPlan = NULL,
            FechaFinPlan = NULL,
            TipoCobro = NULL,
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
            @NegocioId, @SuscripcionId, N'EXTENSION_PRUEBA',
            @EstadoAnterior, 1,
            @EsPruebaAnterior, 1,
            NULL, NULL,
            @FechaInicioPrueba, @NuevaFechaFinPrueba,
            NULL, @DiasExtra, NULLIF(LTRIM(RTRIM(COALESCE(@Observacion, N''))), N''),
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
