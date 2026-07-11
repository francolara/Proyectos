-- =============================================
-- Author:        FRANCO LARA / Codex
-- Create date:   10/07/2026
-- Description:   Crea o actualiza la configuracion operativa y de facturacion de la cuenta administradora.
-- =============================================

CREATE OR ALTER PROCEDURE dbo.usp_SEG_GuardarConfiguracionCuentaAdministradora
    @IdCuentaAdministradora INT,
    @NombreResponsablePrincipal NVARCHAR(180) = NULL,
    @CorreoAdministrativo NVARCHAR(256) = NULL,
    @TelefonoAdministrativo NVARCHAR(30) = NULL,
    @IdEmpresaPredeterminada INT = NULL,
    @ObservacionAdministrativa NVARCHAR(400) = NULL,
    @TipoComprobantePreferido VARCHAR(20) = 'BOLETA',
    @TipoDocumentoFacturacion VARCHAR(20) = 'DNI',
    @NumeroDocumento VARCHAR(20) = NULL,
    @NombreFacturacion NVARCHAR(200) = NULL,
    @RazonSocialFacturacion NVARCHAR(200) = NULL,
    @CorreoFacturacion NVARCHAR(256) = NULL,
    @TelefonoFacturacion NVARCHAR(30) = NULL,
    @DireccionFiscal NVARCHAR(250) = NULL,
    @Ubigeo VARCHAR(6) = NULL,
    @Distrito NVARCHAR(100) = NULL,
    @Provincia NVARCHAR(100) = NULL,
    @Departamento NVARCHAR(100) = NULL,
    @ObservacionFacturacion NVARCHAR(400) = NULL,
    @UsuarioRegistro NVARCHAR(450) = NULL
AS
BEGIN

    SET NOCOUNT ON;

    BEGIN TRY

        IF NOT EXISTS
        (
            SELECT 1
            FROM dbo.SEG_CuentaAdministradora AS ca
            WHERE ca.IdCuentaAdministradora = @IdCuentaAdministradora
              AND ca.Estado = 1
        )
        BEGIN
            RAISERROR (N'La cuenta administradora no existe o no esta activa.', 16, 1);
            RETURN;
        END;

        IF @IdEmpresaPredeterminada IS NOT NULL
           AND NOT EXISTS
        (
            SELECT 1
            FROM dbo.SEG_Empresa AS e
            WHERE e.IdEmpresa = @IdEmpresaPredeterminada
              AND e.IdCuentaAdministradora = @IdCuentaAdministradora
              AND e.Estado = 1
        )
        BEGIN
            RAISERROR (N'La empresa predeterminada no pertenece a la cuenta administradora.', 16, 1);
            RETURN;
        END;

        SET @TipoComprobantePreferido = UPPER(LTRIM(RTRIM(ISNULL(@TipoComprobantePreferido, 'BOLETA'))));
        SET @TipoDocumentoFacturacion = UPPER(LTRIM(RTRIM(ISNULL(@TipoDocumentoFacturacion, 'DNI'))));

        IF @TipoComprobantePreferido NOT IN ('BOLETA', 'FACTURA')
        BEGIN
            RAISERROR (N'El tipo de comprobante preferido no es valido.', 16, 1);
            RETURN;
        END;

        IF @TipoDocumentoFacturacion NOT IN ('DNI', 'RUC', 'CE', 'PASAPORTE', 'OTRO')
        BEGIN
            RAISERROR (N'El tipo de documento de facturacion no es valido.', 16, 1);
            RETURN;
        END;

        IF EXISTS
        (
            SELECT 1
            FROM dbo.SEG_CuentaAdministradoraConfiguracion AS cfg
            WHERE cfg.IdCuentaAdministradora = @IdCuentaAdministradora
        )
        BEGIN
            UPDATE dbo.SEG_CuentaAdministradoraConfiguracion
            SET NombreResponsablePrincipal = @NombreResponsablePrincipal,
                CorreoAdministrativo = @CorreoAdministrativo,
                TelefonoAdministrativo = @TelefonoAdministrativo,
                IdEmpresaPredeterminada = @IdEmpresaPredeterminada,
                ObservacionAdministrativa = @ObservacionAdministrativa,
                Estado = 1,
                FechaActualizacion = SYSDATETIME(),
                UsuarioActualizacion = @UsuarioRegistro
            WHERE IdCuentaAdministradora = @IdCuentaAdministradora;
        END
        ELSE
        BEGIN
            INSERT INTO dbo.SEG_CuentaAdministradoraConfiguracion
            (
                IdCuentaAdministradora,
                NombreResponsablePrincipal,
                CorreoAdministrativo,
                TelefonoAdministrativo,
                IdEmpresaPredeterminada,
                ObservacionAdministrativa,
                Estado,
                UsuarioRegistro
            )
            VALUES
            (
                @IdCuentaAdministradora,
                @NombreResponsablePrincipal,
                @CorreoAdministrativo,
                @TelefonoAdministrativo,
                @IdEmpresaPredeterminada,
                @ObservacionAdministrativa,
                1,
                @UsuarioRegistro
            );
        END;

        IF EXISTS
        (
            SELECT 1
            FROM dbo.SEG_CuentaAdministradoraFacturacion AS fac
            WHERE fac.IdCuentaAdministradora = @IdCuentaAdministradora
        )
        BEGIN
            UPDATE dbo.SEG_CuentaAdministradoraFacturacion
            SET TipoComprobantePreferido = @TipoComprobantePreferido,
                TipoDocumentoFacturacion = @TipoDocumentoFacturacion,
                NumeroDocumento = @NumeroDocumento,
                NombreFacturacion = @NombreFacturacion,
                RazonSocialFacturacion = @RazonSocialFacturacion,
                CorreoFacturacion = @CorreoFacturacion,
                TelefonoFacturacion = @TelefonoFacturacion,
                DireccionFiscal = @DireccionFiscal,
                Ubigeo = @Ubigeo,
                Distrito = @Distrito,
                Provincia = @Provincia,
                Departamento = @Departamento,
                ObservacionFacturacion = @ObservacionFacturacion,
                Estado = 1,
                FechaActualizacion = SYSDATETIME(),
                UsuarioActualizacion = @UsuarioRegistro
            WHERE IdCuentaAdministradora = @IdCuentaAdministradora;
        END
        ELSE
        BEGIN
            INSERT INTO dbo.SEG_CuentaAdministradoraFacturacion
            (
                IdCuentaAdministradora,
                TipoComprobantePreferido,
                TipoDocumentoFacturacion,
                NumeroDocumento,
                NombreFacturacion,
                RazonSocialFacturacion,
                CorreoFacturacion,
                TelefonoFacturacion,
                DireccionFiscal,
                Ubigeo,
                Distrito,
                Provincia,
                Departamento,
                ObservacionFacturacion,
                Estado,
                UsuarioRegistro
            )
            VALUES
            (
                @IdCuentaAdministradora,
                @TipoComprobantePreferido,
                @TipoDocumentoFacturacion,
                @NumeroDocumento,
                @NombreFacturacion,
                @RazonSocialFacturacion,
                @CorreoFacturacion,
                @TelefonoFacturacion,
                @DireccionFiscal,
                @Ubigeo,
                @Distrito,
                @Provincia,
                @Departamento,
                @ObservacionFacturacion,
                1,
                @UsuarioRegistro
            );
        END;

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
