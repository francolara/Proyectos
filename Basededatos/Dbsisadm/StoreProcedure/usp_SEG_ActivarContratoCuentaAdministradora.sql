-- =============================================
-- Author:        FRANCO LARA
-- Create date:   10/07/2026
-- Description:   Activa el contrato comercial de una cuenta administradora y registra el movimiento historico.
-- =============================================
-- =============================================
-- Author:        FRANCO LARA
-- Create date:   25/07/2026
-- Description:   Selecciona BASICO o PRO al iniciar el contrato y aplica los limites comerciales del plan.
-- =============================================

CREATE OR ALTER PROCEDURE dbo.usp_SEG_ActivarContratoCuentaAdministradora
    @IdCuentaAdministradora INT,
    @TipoPlan NVARCHAR(50),
    @TipoCobro NVARCHAR(20),
    @FechaInicioPlan DATE,
    @FechaFinPlan DATE,
    @DiasGracia INT = 5,
    @Observacion NVARCHAR(500) = NULL,
    @UsuarioRegistro NVARCHAR(450) = NULL
AS
BEGIN

    SET NOCOUNT ON;

    BEGIN TRY

        DECLARE @IdCuentaAdministradoraSuscripcion INT;
        DECLARE @TipoPlanAnterior NVARCHAR(50);
        DECLARE @EstadoSuscripcionAnterior NVARCHAR(20);
        DECLARE @EsPruebaAnterior BIT;
        DECLARE @TipoCobroAnterior NVARCHAR(20);
        DECLARE @EmpresasPermitidas INT;
        DECLARE @UsuariosPermitidos INT;
        DECLARE @DiasGraciaNormalizado INT = CASE WHEN @DiasGracia < 0 THEN 0 ELSE @DiasGracia END;
        DECLARE @TipoPlanNormalizado NVARCHAR(50) = UPPER(LTRIM(RTRIM(@TipoPlan)));
        DECLARE @EmpresasPermitidasNuevo INT;
        DECLARE @UsuariosPermitidosNuevo INT;

        IF @TipoPlanNormalizado NOT IN (N'BASICO', N'PRO')
        BEGIN
            RAISERROR (N'El plan debe ser Emprendedor o Contador.', 16, 1);
            RETURN;
        END;

        SET @EmpresasPermitidasNuevo = CASE WHEN @TipoPlanNormalizado = N'PRO' THEN 10 ELSE 3 END;
        SET @UsuariosPermitidosNuevo = CASE WHEN @TipoPlanNormalizado = N'PRO' THEN 3 ELSE 2 END;

        IF @FechaFinPlan < @FechaInicioPlan
        BEGIN
            RAISERROR (N'La fecha fin del contrato no puede ser menor a la fecha inicio.', 16, 1);
            RETURN;
        END;

        SELECT
            @IdCuentaAdministradoraSuscripcion = cas.IdCuentaAdministradoraSuscripcion,
            @TipoPlanAnterior = cas.TipoPlan,
            @EstadoSuscripcionAnterior = cas.EstadoSuscripcion,
            @EsPruebaAnterior = cas.EsPrueba,
            @TipoCobroAnterior = cas.TipoCobro,
            @EmpresasPermitidas = cas.EmpresasPermitidas,
            @UsuariosPermitidos = cas.UsuariosPermitidos
        FROM dbo.SEG_CuentaAdministradoraSuscripcion AS cas
        WHERE cas.IdCuentaAdministradora = @IdCuentaAdministradora;

        IF @IdCuentaAdministradoraSuscripcion IS NULL
        BEGIN
            INSERT INTO dbo.SEG_CuentaAdministradoraSuscripcion
            (
                IdCuentaAdministradora,
                TipoPlan,
                EstadoSuscripcion,
                EsPrueba,
                FechaInicioPlan,
                FechaFinPlan,
                TipoCobro,
                DiasGracia,
                FechaFinGracia,
                EmpresasPermitidas,
                UsuariosPermitidos,
                Activo,
                Observacion,
                UsuarioRegistro,
                FechaActualizacion,
                UsuarioActualizacion
            )
            VALUES
            (
                @IdCuentaAdministradora,
                @TipoPlanNormalizado,
                N'ACTIVO',
                0,
                @FechaInicioPlan,
                @FechaFinPlan,
                @TipoCobro,
                @DiasGraciaNormalizado,
                DATEADD(DAY, @DiasGraciaNormalizado, @FechaFinPlan),
                @EmpresasPermitidasNuevo,
                @UsuariosPermitidosNuevo,
                1,
                @Observacion,
                @UsuarioRegistro,
                SYSDATETIME(),
                @UsuarioRegistro
            );

            SET @IdCuentaAdministradoraSuscripcion = SCOPE_IDENTITY();
            SET @EstadoSuscripcionAnterior = NULL;
            SET @EsPruebaAnterior = NULL;
            SET @TipoCobroAnterior = NULL;
        END
        ELSE
        BEGIN
            UPDATE dbo.SEG_CuentaAdministradoraSuscripcion
            SET TipoPlan = @TipoPlanNormalizado,
                EstadoSuscripcion = N'ACTIVO',
                EsPrueba = 0,
                FechaInicioPrueba = NULL,
                FechaFinPrueba = NULL,
                FechaInicioPlan = @FechaInicioPlan,
                FechaFinPlan = @FechaFinPlan,
                TipoCobro = @TipoCobro,
                DiasGracia = @DiasGraciaNormalizado,
                FechaFinGracia = DATEADD(DAY, @DiasGraciaNormalizado, @FechaFinPlan),
                EmpresasPermitidas = @EmpresasPermitidasNuevo,
                UsuariosPermitidos = @UsuariosPermitidosNuevo,
                Activo = 1,
                Observacion = @Observacion,
                FechaActualizacion = SYSDATETIME(),
                UsuarioActualizacion = @UsuarioRegistro
            WHERE IdCuentaAdministradoraSuscripcion = @IdCuentaAdministradoraSuscripcion;
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
            N'ACTIVACION_CONTRATO',
            @TipoPlanAnterior,
            @TipoPlanNormalizado,
            @EstadoSuscripcionAnterior,
            N'ACTIVO',
            @EsPruebaAnterior,
            0,
            @TipoCobroAnterior,
            @TipoCobro,
            @FechaInicioPlan,
            @FechaFinPlan,
            @DiasGraciaNormalizado,
            0,
            @EmpresasPermitidas,
            @EmpresasPermitidasNuevo,
            @UsuariosPermitidos,
            @UsuariosPermitidosNuevo,
            @Observacion,
            @UsuarioRegistro
        );

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
