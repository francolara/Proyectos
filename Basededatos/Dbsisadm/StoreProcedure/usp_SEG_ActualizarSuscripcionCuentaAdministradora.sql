-- =============================================
-- Author:        FRANCO LARA
-- Create date:   15/06/2026
-- Description:   Actualiza la suscripcion comercial y el estado de la cuenta administradora.
-- =============================================
-- =============================================
-- Author:        FRANCO LARA
-- Create date:   10/07/2026
-- Description:   Integra tipo de cobro y dias de gracia en la actualizacion manual de suscripcion por cuenta.
-- =============================================

CREATE OR ALTER PROCEDURE dbo.usp_SEG_ActualizarSuscripcionCuentaAdministradora
    @IdCuentaAdministradora INT,
    @TipoPlan NVARCHAR(50),
    @EstadoSuscripcion NVARCHAR(20),
    @EsPrueba BIT,
    @FechaInicioPrueba DATE = NULL,
    @FechaFinPrueba DATE = NULL,
    @FechaInicioPlan DATE = NULL,
    @FechaFinPlan DATE = NULL,
    @TipoCobro NVARCHAR(20) = NULL,
    @DiasGracia INT = 5,
    @EmpresasPermitidas INT = NULL,
    @UsuariosPermitidos INT = NULL,
    @Activo BIT,
    @EstadoCuenta BIT,
    @Observacion NVARCHAR(500) = NULL,
    @UsuarioRegistro NVARCHAR(450) = NULL
AS
BEGIN

    SET NOCOUNT ON;

    BEGIN TRY

        DECLARE @IdCuentaAdministradoraSuscripcion INT
        DECLARE @TipoPlanAnterior NVARCHAR(50)
        DECLARE @EstadoSuscripcionAnterior NVARCHAR(20)
        DECLARE @EsPruebaAnterior BIT
        DECLARE @TipoCobroAnterior NVARCHAR(20)
        DECLARE @EmpresasPermitidasAnterior INT
        DECLARE @UsuariosPermitidosAnterior INT

        UPDATE dbo.SEG_CuentaAdministradora
        SET Estado = @EstadoCuenta,
            UsuarioRegistro = @UsuarioRegistro
        WHERE IdCuentaAdministradora = @IdCuentaAdministradora;

        SELECT
            @IdCuentaAdministradoraSuscripcion = cas.IdCuentaAdministradoraSuscripcion,
            @TipoPlanAnterior = cas.TipoPlan,
            @EstadoSuscripcionAnterior = cas.EstadoSuscripcion,
            @EsPruebaAnterior = cas.EsPrueba,
            @TipoCobroAnterior = cas.TipoCobro,
            @EmpresasPermitidasAnterior = cas.EmpresasPermitidas,
            @UsuariosPermitidosAnterior = cas.UsuariosPermitidos
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
                FechaInicioPrueba,
                FechaFinPrueba,
                FechaInicioPlan,
                FechaFinPlan,
                TipoCobro,
                DiasGracia,
                FechaFinGracia,
                EmpresasPermitidas,
                UsuariosPermitidos,
                Activo,
                Observacion,
                UsuarioRegistro
            )
            VALUES
            (
                @IdCuentaAdministradora,
                @TipoPlan,
                @EstadoSuscripcion,
                @EsPrueba,
                @FechaInicioPrueba,
                @FechaFinPrueba,
                @FechaInicioPlan,
                @FechaFinPlan,
                @TipoCobro,
                CASE WHEN @DiasGracia < 0 THEN 0 ELSE @DiasGracia END,
                CASE
                    WHEN @FechaFinPlan IS NULL THEN NULL
                    ELSE DATEADD(DAY, CASE WHEN @DiasGracia < 0 THEN 0 ELSE @DiasGracia END, @FechaFinPlan)
                END,
                @EmpresasPermitidas,
                @UsuariosPermitidos,
                @Activo,
                @Observacion,
                @UsuarioRegistro
            );

            SET @IdCuentaAdministradoraSuscripcion = SCOPE_IDENTITY();
        END
        ELSE
        BEGIN
            UPDATE dbo.SEG_CuentaAdministradoraSuscripcion
            SET TipoPlan = @TipoPlan,
                EstadoSuscripcion = @EstadoSuscripcion,
                EsPrueba = @EsPrueba,
                FechaInicioPrueba = @FechaInicioPrueba,
                FechaFinPrueba = @FechaFinPrueba,
                FechaInicioPlan = @FechaInicioPlan,
                FechaFinPlan = @FechaFinPlan,
                TipoCobro = @TipoCobro,
                DiasGracia = CASE WHEN @DiasGracia < 0 THEN 0 ELSE @DiasGracia END,
                FechaFinGracia = CASE
                    WHEN @FechaFinPlan IS NULL THEN NULL
                    ELSE DATEADD(DAY, CASE WHEN @DiasGracia < 0 THEN 0 ELSE @DiasGracia END, @FechaFinPlan)
                END,
                EmpresasPermitidas = @EmpresasPermitidas,
                UsuariosPermitidos = @UsuariosPermitidos,
                Activo = @Activo,
                Observacion = @Observacion,
                UsuarioRegistro = @UsuarioRegistro,
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
            N'ACTUALIZACION_MANUAL',
            @TipoPlanAnterior,
            @TipoPlan,
            @EstadoSuscripcionAnterior,
            @EstadoSuscripcion,
            @EsPruebaAnterior,
            @EsPrueba,
            @TipoCobroAnterior,
            @TipoCobro,
            COALESCE(@FechaInicioPlan, @FechaInicioPrueba),
            COALESCE(@FechaFinPlan, @FechaFinPrueba),
            CASE WHEN @DiasGracia < 0 THEN 0 ELSE @DiasGracia END,
            0,
            @EmpresasPermitidasAnterior,
            @EmpresasPermitidas,
            @UsuariosPermitidosAnterior,
            @UsuariosPermitidos,
            @Observacion,
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
