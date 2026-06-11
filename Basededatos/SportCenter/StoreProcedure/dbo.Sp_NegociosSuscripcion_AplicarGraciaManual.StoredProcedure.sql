-- =============================================
-- Author:        FRANCO LARA
-- Create date:   10/06/2026
-- Firma:         Aplica dias de gracia manual a un contrato y registra movimiento comercial.
-- =============================================
CREATE OR ALTER PROCEDURE dbo.Sp_NegociosSuscripcion_AplicarGraciaManual
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
        DECLARE @TipoCobroAnterior NVARCHAR(20);
        DECLARE @FechaInicioPlan DATE;
        DECLARE @FechaFinPlan DATE;
        DECLARE @DiasGraciaActual INT;
        DECLARE @FechaFinGraciaActual DATE;
        DECLARE @NuevaFechaFinGracia DATE;
        DECLARE @Hoy DATE = CAST(GETDATE() AS DATE);

        IF @NegocioId IS NULL OR @NegocioId <= 0
            RAISERROR('Negocio invalido.', 16, 1);

        IF @DiasExtra IS NULL OR @DiasExtra <= 0
            RAISERROR('Ingresa la cantidad de dias de gracia a agregar.', 16, 1);

        IF NOT EXISTS (SELECT 1 FROM dbo.Negocios WHERE Id = @NegocioId)
            RAISERROR('Negocio no encontrado.', 16, 1);

        SELECT
            @SuscripcionId = ns.Id,
            @EstadoAnterior = ns.EstadoSuscripcion,
            @EsPruebaAnterior = ns.EsPrueba,
            @TipoCobroAnterior = UPPER(LTRIM(RTRIM(COALESCE(ns.TipoCobro, N'')))),
            @FechaInicioPlan = ns.FechaInicioPlan,
            @FechaFinPlan = ns.FechaFinPlan,
            @DiasGraciaActual = COALESCE(ns.DiasGracia, 0),
            @FechaFinGraciaActual = ns.FechaFinGracia
        FROM dbo.NegociosSuscripcion ns
        WHERE ns.NegocioId = @NegocioId;

        IF @SuscripcionId IS NULL
            RAISERROR('El negocio no tiene suscripcion creada.', 16, 1);

        IF COALESCE(@EsPruebaAnterior, 0) = 1 OR NULLIF(@TipoCobroAnterior, N'') IS NULL
            RAISERROR('La gracia manual solo aplica a contratos comerciales.', 16, 1);

        SET @NuevaFechaFinGracia = DATEADD(DAY, @DiasExtra, CASE
            WHEN @FechaFinGraciaActual IS NULL AND @FechaFinPlan IS NOT NULL AND @FechaFinPlan >= @Hoy THEN @FechaFinPlan
            WHEN @FechaFinGraciaActual IS NULL AND @FechaFinPlan IS NOT NULL THEN @Hoy
            WHEN @FechaFinGraciaActual < @Hoy THEN @Hoy
            ELSE @FechaFinGraciaActual
        END);

        UPDATE dbo.NegociosSuscripcion
        SET EstadoSuscripcion = 2,
            EsPrueba = 0,
            FechaFinGracia = @NuevaFechaFinGracia,
            DiasGracia = COALESCE(@DiasGraciaActual, 0) + @DiasExtra,
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
            @NegocioId, @SuscripcionId, N'GRACIA_MANUAL',
            @EstadoAnterior, 2,
            @EsPruebaAnterior, 0,
            @TipoCobroAnterior, @TipoCobroAnterior,
            @FechaInicioPlan, @NuevaFechaFinGracia,
            COALESCE(@DiasGraciaActual, 0) + @DiasExtra, @DiasExtra,
            NULLIF(LTRIM(RTRIM(COALESCE(@Observacion, N''))), N''),
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
