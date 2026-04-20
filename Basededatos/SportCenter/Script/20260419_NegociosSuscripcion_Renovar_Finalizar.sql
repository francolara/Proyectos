-- =============================================
-- Author:        FRANCO LARA
-- Create date:   19/04/2026
-- Firma:         Despliegue de procedimientos para renovar y finalizar contratos de suscripcion de negocios.
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
        DECLARE @Hoy DATE = CAST(GETDATE() AS DATE);

        IF NOT EXISTS (SELECT 1 FROM dbo.Negocios WHERE Id = @NegocioId)
            RAISERROR('Negocio no encontrado.', 16, 1);

        SELECT
            @TipoCobro = UPPER(LTRIM(RTRIM(COALESCE(ns.TipoCobro, N'')))),
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
