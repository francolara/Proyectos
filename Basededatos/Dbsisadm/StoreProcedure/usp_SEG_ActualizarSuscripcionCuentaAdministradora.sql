-- =============================================
-- Author:        FRANCO LARA
-- Create date:   15/06/2026
-- Description:   Actualiza la suscripcion comercial y el estado de la cuenta administradora.
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
                EmpresasPermitidas = @EmpresasPermitidas,
                UsuariosPermitidos = @UsuariosPermitidos,
                Activo = @Activo,
                Observacion = @Observacion,
                UsuarioRegistro = @UsuarioRegistro
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
            FechaInicioReferencia,
            FechaFinReferencia,
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
            COALESCE(@FechaInicioPlan, @FechaInicioPrueba),
            COALESCE(@FechaFinPlan, @FechaFinPrueba),
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
