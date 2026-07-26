-- =============================================
-- Author:        FRANCO LARA
-- Create date:   25/07/2026
-- Description:   Persiste como suspendida una suscripcion vencida al evaluar el acceso de la cuenta.
-- =============================================

CREATE OR ALTER PROCEDURE dbo.usp_SEG_SincronizarVencimientoSuscripcionCuentaAdministradora
    @IdCuentaAdministradora INT,
    @UsuarioRegistro NVARCHAR(450) = NULL
AS
BEGIN

    SET NOCOUNT ON;

    BEGIN TRY

        DECLARE @IdCuentaAdministradoraSuscripcion INT;
        DECLARE @TipoPlan NVARCHAR(50);
        DECLARE @EstadoSuscripcion NVARCHAR(20);
        DECLARE @EsPrueba BIT;
        DECLARE @FechaInicioPrueba DATE;
        DECLARE @FechaFinPrueba DATE;
        DECLARE @FechaInicioPlan DATE;
        DECLARE @FechaFinPlan DATE;
        DECLARE @FechaFinGracia DATE;
        DECLARE @TipoCobro NVARCHAR(20);
        DECLARE @DiasGracia INT;
        DECLARE @EmpresasPermitidas INT;
        DECLARE @UsuariosPermitidos INT;
        DECLARE @FechaActual DATE = CAST(GETDATE() AS DATE);
        DECLARE @DebeSuspender BIT = 0;

        SELECT
            @IdCuentaAdministradoraSuscripcion = cas.IdCuentaAdministradoraSuscripcion,
            @TipoPlan = cas.TipoPlan,
            @EstadoSuscripcion = cas.EstadoSuscripcion,
            @EsPrueba = cas.EsPrueba,
            @FechaInicioPrueba = cas.FechaInicioPrueba,
            @FechaFinPrueba = cas.FechaFinPrueba,
            @FechaInicioPlan = cas.FechaInicioPlan,
            @FechaFinPlan = cas.FechaFinPlan,
            @FechaFinGracia = cas.FechaFinGracia,
            @TipoCobro = cas.TipoCobro,
            @DiasGracia = cas.DiasGracia,
            @EmpresasPermitidas = cas.EmpresasPermitidas,
            @UsuariosPermitidos = cas.UsuariosPermitidos
        FROM dbo.SEG_CuentaAdministradoraSuscripcion AS cas
        WHERE cas.IdCuentaAdministradora = @IdCuentaAdministradora;

        IF @IdCuentaAdministradoraSuscripcion IS NULL
           OR UPPER(ISNULL(@EstadoSuscripcion, N'')) IN (N'SUSPENDIDO', N'BAJA')
        BEGIN
            RETURN;
        END;

        IF
        (
            ISNULL(@EsPrueba, 0) = 1
            OR UPPER(ISNULL(@TipoPlan, N'')) IN (N'TRIAL', N'GRATIS')
            OR UPPER(ISNULL(@EstadoSuscripcion, N'')) = N'TRIAL'
        )
        AND @FechaFinPrueba IS NOT NULL
        AND @FechaActual > @FechaFinPrueba
        BEGIN
            SET @DebeSuspender = 1;
        END;

        IF
        (
            ISNULL(@EsPrueba, 0) = 0
            AND UPPER(ISNULL(@TipoPlan, N'')) NOT IN (N'TRIAL', N'GRATIS')
            AND UPPER(ISNULL(@EstadoSuscripcion, N'')) <> N'TRIAL'
        )
        AND @FechaFinPlan IS NOT NULL
        AND @FechaActual >
            COALESCE
            (
                @FechaFinGracia,
                DATEADD(DAY, CASE WHEN ISNULL(@DiasGracia, 0) < 0 THEN 0 ELSE ISNULL(@DiasGracia, 0) END, @FechaFinPlan)
            )
        BEGIN
            SET @DebeSuspender = 1;
        END;

        IF @DebeSuspender = 0
        BEGIN
            RETURN;
        END;

        UPDATE dbo.SEG_CuentaAdministradoraSuscripcion
        SET EstadoSuscripcion = N'SUSPENDIDO',
            Activo = 0,
            FechaActualizacion = SYSDATETIME(),
            UsuarioActualizacion = @UsuarioRegistro
        WHERE IdCuentaAdministradoraSuscripcion = @IdCuentaAdministradoraSuscripcion
          AND UPPER(ISNULL(EstadoSuscripcion, N'')) NOT IN (N'SUSPENDIDO', N'BAJA');

        IF @@ROWCOUNT = 0
        BEGIN
            RETURN;
        END;

        INSERT INTO dbo.SEG_CuentaAdministradoraSuscripcionMovimiento
        (
            IdCuentaAdministradora,
            IdCuentaAdministradoraSuscripcion,
            TipoMovimiento,
            TipoPlanAnterior,
            TipoPlanNuevo,
            EstadoSuscripcionAnterior,
            EstadoSuscripcionNuevo,
            EsPruebaAnterior,
            EsPruebaNuevo,
            TipoCobroAnterior,
            TipoCobroNuevo,
            FechaInicioReferencia,
            FechaFinReferencia,
            DiasGracia,
            DiasExtra,
            EmpresasPermitidasAnterior,
            EmpresasPermitidasNuevo,
            UsuariosPermitidosAnterior,
            UsuariosPermitidosNuevo,
            Observacion,
            UsuarioRegistro
        )
        VALUES
        (
            @IdCuentaAdministradora,
            @IdCuentaAdministradoraSuscripcion,
            N'SUSPENSION_VENCIMIENTO',
            @TipoPlan,
            @TipoPlan,
            @EstadoSuscripcion,
            N'SUSPENDIDO',
            @EsPrueba,
            @EsPrueba,
            @TipoCobro,
            @TipoCobro,
            CASE
                WHEN ISNULL(@EsPrueba, 0) = 1 OR UPPER(ISNULL(@TipoPlan, N'')) IN (N'TRIAL', N'GRATIS')
                    THEN @FechaInicioPrueba
                ELSE @FechaInicioPlan
            END,
            CASE
                WHEN ISNULL(@EsPrueba, 0) = 1 OR UPPER(ISNULL(@TipoPlan, N'')) IN (N'TRIAL', N'GRATIS')
                    THEN @FechaFinPrueba
                ELSE COALESCE(@FechaFinGracia, @FechaFinPlan)
            END,
            ISNULL(@DiasGracia, 0),
            0,
            @EmpresasPermitidas,
            @EmpresasPermitidas,
            @UsuariosPermitidos,
            @UsuariosPermitidos,
            N'Suspension automatica por vencimiento al evaluar el acceso.',
            @UsuarioRegistro
        );

    END TRY

    BEGIN CATCH

        DECLARE @ErrorMessage NVARCHAR(4000)
        DECLARE @ErrorSeverity INT
        DECLARE @ErrorState INT

        SELECT
            @ErrorMessage = ERROR_MESSAGE(),
            @ErrorSeverity = ERROR_SEVERITY(),
            @ErrorState = ERROR_STATE()

        RAISERROR (@ErrorMessage, @ErrorSeverity, @ErrorState)

    END CATCH

END
