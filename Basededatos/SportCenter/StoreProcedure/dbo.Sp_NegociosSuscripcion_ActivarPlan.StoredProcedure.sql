-- =============================================
-- Author:        FRANCO LARA
-- Create date:   16/04/2026
-- Firma:         Activacion de contrato con tipo de cobro, vigencia y periodo de gracia.
-- Firma:         10/06/2026 | Registra movimiento comercial de activacion o reactivacion para trazabilidad del contrato.
-- =============================================
CREATE OR ALTER PROCEDURE dbo.Sp_NegociosSuscripcion_ActivarPlan
    @NegocioId INT,
    @TipoCobro NVARCHAR(20),
    @FechaInicioPlan DATE,
    @FechaFinPlan DATE,
    @DiasGracia INT = 5,
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

        IF @TipoCobroNorm NOT IN (N'MENSUAL', N'TRIMESTRAL', N'SEMESTRAL', N'ANUAL')
            RAISERROR('Tipo de cobro invalido. Usa MENSUAL, TRIMESTRAL, SEMESTRAL o ANUAL.', 16, 1);

        IF @FechaInicioPlan IS NULL OR @FechaFinPlan IS NULL OR @FechaFinPlan < @FechaInicioPlan
            RAISERROR('Rango de vigencia invalido para el plan.', 16, 1);

        IF @DiasGracia IS NULL OR @DiasGracia < 0
            SET @DiasGracia = 5;

        IF NOT EXISTS (SELECT 1 FROM dbo.Negocios WHERE Id = @NegocioId)
            RAISERROR('Negocio no encontrado.', 16, 1);

        IF OBJECT_ID(N'dbo.NegociosSuscripcion', N'U') IS NULL
            RAISERROR('No existe la tabla NegociosSuscripcion.', 16, 1);

        SELECT
            @SuscripcionId = ns.Id,
            @EstadoAnterior = ns.EstadoSuscripcion,
            @EsPruebaAnterior = ns.EsPrueba,
            @TipoCobroAnterior = UPPER(LTRIM(RTRIM(COALESCE(ns.TipoCobro, N''))))
        FROM dbo.NegociosSuscripcion ns
        WHERE ns.NegocioId = @NegocioId;

        IF EXISTS (SELECT 1 FROM dbo.NegociosSuscripcion WHERE NegocioId = @NegocioId)
        BEGIN
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

            SELECT @SuscripcionId = ns.Id
            FROM dbo.NegociosSuscripcion ns
            WHERE ns.NegocioId = @NegocioId;
        END
        ELSE
        BEGIN
            INSERT INTO dbo.NegociosSuscripcion
            (
                NegocioId, EstadoSuscripcion, EsPrueba,
                FechaInicioPrueba, FechaFinPrueba,
                FechaInicioPlan, FechaFinPlan,
                TipoCobro, DiasGracia, FechaFinGracia,
                FechaCreacion, UsuarioCreacion
            )
            VALUES
            (
                @NegocioId, 2, 0,
                NULL, NULL,
                @FechaInicioPlan, @FechaFinPlan,
                @TipoCobroNorm, @DiasGracia, DATEADD(DAY, @DiasGracia, @FechaFinPlan),
                SYSUTCDATETIME(), COALESCE(@Usuario, N'sistema')
            );

            SET @SuscripcionId = CAST(SCOPE_IDENTITY() AS INT);
        END;

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
                @NegocioId, @SuscripcionId, N'ACTIVACION_CONTRATO',
                @EstadoAnterior, 2,
                @EsPruebaAnterior, 0,
                NULLIF(@TipoCobroAnterior, N''), @TipoCobroNorm,
                @FechaInicioPlan, @FechaFinPlan,
                @DiasGracia, NULL, N'Inicio o reactivacion manual de contrato desde superadmin.',
                SYSUTCDATETIME(), COALESCE(@Usuario, N'sistema')
            );
        END;
    END TRY
    BEGIN CATCH
        DECLARE @ErrorMessage NVARCHAR(4000), @ErrorSeverity INT, @ErrorState INT;
        SELECT @ErrorMessage = ERROR_MESSAGE(), @ErrorSeverity = ERROR_SEVERITY(), @ErrorState = ERROR_STATE();
        RAISERROR (@ErrorMessage, @ErrorSeverity, @ErrorState);
    END CATCH
END
