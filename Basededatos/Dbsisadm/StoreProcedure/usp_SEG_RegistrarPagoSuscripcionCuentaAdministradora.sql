-- =============================================
-- Author:        FRANCO LARA
-- Create date:   10/07/2026
-- Description:   Registra cobros manuales o conciliables de la suscripcion por cuenta administradora.
-- =============================================

CREATE OR ALTER PROCEDURE dbo.usp_SEG_RegistrarPagoSuscripcionCuentaAdministradora
    @IdCuentaAdministradora INT,
    @TipoPago NVARCHAR(30),
    @EstadoPago NVARCHAR(20),
    @Monto DECIMAL(12,2),
    @Moneda NVARCHAR(10) = N'PEN',
    @FechaPago DATETIME2(0),
    @FechaVencimiento DATE = NULL,
    @OperacionNumero NVARCHAR(100) = NULL,
    @EntidadFinanciera NVARCHAR(120) = NULL,
    @ReferenciaExterna NVARCHAR(120) = NULL,
    @ProveedorPasarela NVARCHAR(50) = NULL,
    @TransaccionPasarelaId NVARCHAR(120) = NULL,
    @PagoPasarelaId NVARCHAR(120) = NULL,
    @EstadoPasarela NVARCHAR(30) = NULL,
    @PayloadPasarela NVARCHAR(MAX) = NULL,
    @Observacion NVARCHAR(500) = NULL,
    @AccionAplicacion NVARCHAR(30) = NULL,
    @AplicarAlConfirmar BIT = 0,
    @TipoCobroObjetivo NVARCHAR(20) = NULL,
    @FechaInicioPlanObjetivo DATE = NULL,
    @DiasGraciaObjetivo INT = NULL,
    @UsuarioRegistro NVARCHAR(450) = NULL
AS
BEGIN

    SET NOCOUNT ON;

    BEGIN TRY

        DECLARE @IdCuentaAdministradoraSuscripcion INT;
        DECLARE @IdCuentaAdministradoraSuscripcionPago INT;
        DECLARE @DiasGraciaNormalizado INT = CASE WHEN @DiasGraciaObjetivo IS NULL OR @DiasGraciaObjetivo < 0 THEN 5 ELSE @DiasGraciaObjetivo END;
        DECLARE @FechaFinPlanObjetivo DATE;
        DECLARE @TipoCobroContrato NVARCHAR(20);
        DECLARE @ObservacionContrato NVARCHAR(500);

        SELECT
            @IdCuentaAdministradoraSuscripcion = cas.IdCuentaAdministradoraSuscripcion
        FROM dbo.SEG_CuentaAdministradoraSuscripcion AS cas
        WHERE cas.IdCuentaAdministradora = @IdCuentaAdministradora;

        IF @IdCuentaAdministradoraSuscripcion IS NULL
        BEGIN
            RAISERROR (N'La cuenta administradora no tiene una suscripcion creada.', 16, 1);
            RETURN;
        END;

        INSERT INTO dbo.SEG_CuentaAdministradoraSuscripcionPago
        (
            IdCuentaAdministradora,
            IdCuentaAdministradoraSuscripcion,
            TipoPago,
            EstadoPago,
            Monto,
            Moneda,
            FechaPago,
            FechaVencimiento,
            OperacionNumero,
            EntidadFinanciera,
            ReferenciaExterna,
            ProveedorPasarela,
            TransaccionPasarelaId,
            PagoPasarelaId,
            EstadoPasarela,
            PayloadPasarela,
            AccionAplicacion,
            AplicarAlConfirmar,
            AplicadoSuscripcion,
            TipoCobroObjetivo,
            FechaInicioPlanObjetivo,
            DiasGraciaObjetivo,
            Observacion,
            UsuarioRegistro
        )
        VALUES
        (
            @IdCuentaAdministradora,
            @IdCuentaAdministradoraSuscripcion,
            UPPER(LTRIM(RTRIM(@TipoPago))),
            UPPER(LTRIM(RTRIM(@EstadoPago))),
            @Monto,
            UPPER(LTRIM(RTRIM(@Moneda))),
            @FechaPago,
            @FechaVencimiento,
            @OperacionNumero,
            @EntidadFinanciera,
            @ReferenciaExterna,
            @ProveedorPasarela,
            @TransaccionPasarelaId,
            @PagoPasarelaId,
            @EstadoPasarela,
            @PayloadPasarela,
            CASE WHEN NULLIF(LTRIM(RTRIM(@AccionAplicacion)), N'') IS NULL THEN NULL ELSE UPPER(LTRIM(RTRIM(@AccionAplicacion))) END,
            @AplicarAlConfirmar,
            0,
            CASE WHEN NULLIF(LTRIM(RTRIM(@TipoCobroObjetivo)), N'') IS NULL THEN NULL ELSE UPPER(LTRIM(RTRIM(@TipoCobroObjetivo))) END,
            @FechaInicioPlanObjetivo,
            @DiasGraciaObjetivo,
            @Observacion,
            @UsuarioRegistro
        );

        SET @IdCuentaAdministradoraSuscripcionPago = SCOPE_IDENTITY();

        IF UPPER(LTRIM(RTRIM(@EstadoPago))) = N'PAGADO'
           AND @AplicarAlConfirmar = 1
           AND UPPER(LTRIM(RTRIM(ISNULL(@AccionAplicacion, N'')))) = N'ACTIVAR_CONTRATO'
           AND @FechaInicioPlanObjetivo IS NOT NULL
        BEGIN
            SET @TipoCobroContrato = COALESCE(NULLIF(LTRIM(RTRIM(@TipoCobroObjetivo)), N''), N'MENSUAL');
            SET @ObservacionContrato = COALESCE(NULLIF(LTRIM(RTRIM(@Observacion)), N''), N'Cobro aplicado automaticamente al contrato.');

            SET @FechaFinPlanObjetivo =
                CASE UPPER(LTRIM(RTRIM(@TipoCobroContrato)))
                    WHEN N'TRIMESTRAL' THEN DATEADD(DAY, -1, DATEADD(MONTH, 3, @FechaInicioPlanObjetivo))
                    WHEN N'SEMESTRAL' THEN DATEADD(DAY, -1, DATEADD(MONTH, 6, @FechaInicioPlanObjetivo))
                    WHEN N'ANUAL' THEN DATEADD(DAY, -1, DATEADD(YEAR, 1, @FechaInicioPlanObjetivo))
                    ELSE DATEADD(DAY, -1, DATEADD(MONTH, 1, @FechaInicioPlanObjetivo))
                END;

            EXEC dbo.usp_SEG_ActivarContratoCuentaAdministradora
                @IdCuentaAdministradora = @IdCuentaAdministradora,
                @TipoCobro = @TipoCobroContrato,
                @FechaInicioPlan = @FechaInicioPlanObjetivo,
                @FechaFinPlan = @FechaFinPlanObjetivo,
                @DiasGracia = @DiasGraciaNormalizado,
                @Observacion = @ObservacionContrato,
                @UsuarioRegistro = @UsuarioRegistro;

            UPDATE dbo.SEG_CuentaAdministradoraSuscripcionPago
            SET AplicadoSuscripcion = 1,
                FechaAplicacion = SYSDATETIME(),
                UsuarioAplicacion = @UsuarioRegistro,
                FechaActualizacion = SYSDATETIME(),
                UsuarioActualizacion = @UsuarioRegistro
            WHERE IdCuentaAdministradoraSuscripcionPago = @IdCuentaAdministradoraSuscripcionPago;
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
