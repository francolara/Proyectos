-- =============================================
-- Author:        FRANCO LARA
-- Create date:   15/06/2026
-- Description:   Obtiene la empresa activa y la suscripcion de la cuenta administradora a la que pertenece.
-- =============================================

CREATE OR ALTER PROCEDURE dbo.usp_SEG_ObtenerContextoSuscripcionPorEmpresa
    @IdEmpresa INT
AS
BEGIN

    SET NOCOUNT ON;

    BEGIN TRY

        SELECT
            e.IdEmpresa,
            e.IdCuentaAdministradora,
            e.CodigoEmpresa,
            e.RazonSocial,
            e.NombreComercial,
            e.Ruc,
            e.Estado AS EstadoEmpresa,
            ca.CodigoCuenta,
            ca.NombreCuenta,
            ca.CorreoPrincipal,
            ca.TelefonoPrincipal,
            ca.Estado AS EstadoCuenta,
            cas.IdCuentaAdministradoraSuscripcion,
            cas.TipoPlan,
            cas.EstadoSuscripcion,
            cas.EsPrueba,
            cas.FechaInicioPrueba,
            cas.FechaFinPrueba,
            cas.FechaInicioPlan,
            cas.FechaFinPlan,
            cas.EmpresasPermitidas,
            cas.UsuariosPermitidos,
            cas.Activo,
            cas.Observacion
        FROM dbo.SEG_Empresa AS e
        INNER JOIN dbo.SEG_CuentaAdministradora AS ca
            ON ca.IdCuentaAdministradora = e.IdCuentaAdministradora
        LEFT JOIN dbo.SEG_CuentaAdministradoraSuscripcion AS cas
            ON cas.IdCuentaAdministradora = ca.IdCuentaAdministradora
        WHERE e.IdEmpresa = @IdEmpresa;

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
