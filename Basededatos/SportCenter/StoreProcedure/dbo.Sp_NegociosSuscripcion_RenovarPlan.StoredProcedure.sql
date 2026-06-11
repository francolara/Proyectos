-- =============================================
-- Author:        FRANCO LARA
-- Create date:   19/04/2026
-- Firma:         Renovacion automatica de suscripcion segun el tipo de cobro vigente del negocio.
-- Firma:         10/06/2026 | Registra movimiento comercial de renovacion para conservar historial de vigencias.
-- =============================================
CREATE OR ALTER PROCEDURE dbo.Sp_NegociosSuscripcion_RenovarPlan
    @NegocioId INT,
    @Usuario NVARCHAR(200) = NULL
AS
BEGIN
    SET NOCOUNT ON;

    BEGIN TRY
        DECLARE @TipoCobro NVARCHAR(20);
        DECLARE @FechaFinPlanActual DATE;
        DECLARE @DiasGracia INT;
        DECLARE @BaseFechaFin DATE;
        DECLARE @NuevaFechaFin DATE;
        DECLARE @SuscripcionId INT;
        DECLARE @EstadoAnterior INT;
        DECLARE @EsPruebaAnterior BIT;
        DECLARE @FechaInicioPlan DATE;
        DECLARE @Hoy DATE = CAST(GETDATE() AS DATE);

        IF NOT EXISTS (SELECT 1 FROM dbo.Negocios WHERE Id = @NegocioId)
            RAISERROR('Negocio no encontrado.', 16, 1);

        SELECT
            @SuscripcionId = ns.Id,
            @EstadoAnterior = ns.EstadoSuscripcion,
            @EsPruebaAnterior = ns.EsPrueba,
            @TipoCobro = UPPER(LTRIM(RTRIM(COALESCE(ns.TipoCobro, N'')))),
            @FechaInicioPlan = ns.FechaInicioPlan,
            @FechaFinPlanActual = ns.FechaFinPlan,
            @DiasGracia = COALESCE(ns.DiasGracia, 5)
        FROM dbo.NegociosSuscripcion ns
        WHERE ns.NegocioId = @NegocioId;

        IF @TipoCobro NOT IN (N'MENSUAL', N'TRIMESTRAL', N'SEMESTRAL', N'ANUAL')
            RAISERROR('El negocio no tiene un contrato activo renovable.', 16, 1);

        SET @BaseFechaFin = COALESCE(@FechaFinPlanActual, @Hoy);
        IF @BaseFechaFin < @Hoy
            SET @BaseFechaFin = @Hoy;

        SET @NuevaFechaFin =
            CASE @TipoCobro
                WHEN N'MENSUAL' THEN DATEADD(MONTH, 1, @BaseFechaFin)
                WHEN N'TRIMESTRAL' THEN DATEADD(MONTH, 3, @BaseFechaFin)
                WHEN N'SEMESTRAL' THEN DATEADD(MONTH, 6, @BaseFechaFin)
                WHEN N'ANUAL' THEN DATEADD(YEAR, 1, @BaseFechaFin)
            END;

        UPDATE dbo.NegociosSuscripcion
        SET EstadoSuscripcion = 2,
            EsPrueba = 0,
            FechaInicioPrueba = NULL,
            FechaFinPrueba = NULL,
            FechaFinPlan = @NuevaFechaFin,
            FechaFinGracia = DATEADD(DAY, COALESCE(@DiasGracia, 5), @NuevaFechaFin),
            FechaActualizacion = SYSUTCDATETIME(),
            UsuarioActualizacion = COALESCE(@Usuario, N'sistema')
        WHERE NegocioId = @NegocioId;

        IF @@ROWCOUNT = 0
            RAISERROR('No se encontro una suscripcion para renovar.', 16, 1);

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
                @NegocioId, @SuscripcionId, N'RENOVACION',
                @EstadoAnterior, 2,
                @EsPruebaAnterior, 0,
                @TipoCobro, @TipoCobro,
                COALESCE(@FechaInicioPlan, @Hoy), @NuevaFechaFin,
                COALESCE(@DiasGracia, 5), NULL, N'Renovacion desde superadmin segun el tipo de cobro vigente.',
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
