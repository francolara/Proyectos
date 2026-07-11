-- =============================================
-- Author:        FRANCO LARA / Codex
-- Create date:   10/07/2026
-- Description:   Obtiene la configuracion operativa y de facturacion de la cuenta administradora.
-- =============================================

CREATE OR ALTER PROCEDURE dbo.usp_SEG_ObtenerConfiguracionCuentaAdministradora
    @IdCuentaAdministradora INT
AS
BEGIN

    SET NOCOUNT ON;

    BEGIN TRY

        SELECT
            ca.IdCuentaAdministradora,
            ca.CodigoCuenta,
            ca.NombreCuenta,
            ca.CorreoPrincipal,
            ca.TelefonoPrincipal,
            cfg.IdCuentaAdministradoraConfiguracion,
            cfg.NombreResponsablePrincipal,
            cfg.CorreoAdministrativo,
            cfg.TelefonoAdministrativo,
            cfg.IdEmpresaPredeterminada,
            emp.CodigoEmpresa AS CodigoEmpresaPredeterminada,
            emp.RazonSocial AS RazonSocialEmpresaPredeterminada,
            cfg.ObservacionAdministrativa,
            fac.IdCuentaAdministradoraFacturacion,
            fac.TipoComprobantePreferido,
            fac.TipoDocumentoFacturacion,
            fac.NumeroDocumento,
            fac.NombreFacturacion,
            fac.RazonSocialFacturacion,
            fac.CorreoFacturacion,
            fac.TelefonoFacturacion,
            fac.DireccionFiscal,
            fac.Ubigeo,
            fac.Distrito,
            fac.Provincia,
            fac.Departamento,
            fac.ObservacionFacturacion
        FROM dbo.SEG_CuentaAdministradora AS ca
        LEFT JOIN dbo.SEG_CuentaAdministradoraConfiguracion AS cfg
            ON cfg.IdCuentaAdministradora = ca.IdCuentaAdministradora
           AND cfg.Estado = 1
        LEFT JOIN dbo.SEG_Empresa AS emp
            ON emp.IdEmpresa = cfg.IdEmpresaPredeterminada
        LEFT JOIN dbo.SEG_CuentaAdministradoraFacturacion AS fac
            ON fac.IdCuentaAdministradora = ca.IdCuentaAdministradora
           AND fac.Estado = 1
        WHERE ca.IdCuentaAdministradora = @IdCuentaAdministradora
          AND ca.Estado = 1;

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
