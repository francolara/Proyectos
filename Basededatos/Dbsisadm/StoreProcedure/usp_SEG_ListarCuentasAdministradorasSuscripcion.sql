-- =============================================
-- Author:        FRANCO LARA
-- Create date:   15/06/2026
-- Description:   Lista las cuentas administradoras con su estado comercial y empresa principal.
-- =============================================

CREATE OR ALTER PROCEDURE dbo.usp_SEG_ListarCuentasAdministradorasSuscripcion
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
            ca.Estado AS EstadoCuenta,
            principal.IdEmpresa AS IdEmpresaPrincipal,
            principal.CodigoEmpresa AS CodigoEmpresaPrincipal,
            principal.RazonSocial AS RazonSocialEmpresaPrincipal,
            principal.NombreComercial AS NombreComercialEmpresaPrincipal,
            principal.Ruc AS RucEmpresaPrincipal,
            empresas.CantidadEmpresas,
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
            cas.Observacion,
            uca.AspNetUserId,
            up.NombreCompleto,
            up.Telefono,
            au.Email
        FROM dbo.SEG_CuentaAdministradora AS ca
        OUTER APPLY
        (
            SELECT
                COUNT(1) AS CantidadEmpresas
            FROM dbo.SEG_Empresa AS e
            WHERE e.IdCuentaAdministradora = ca.IdCuentaAdministradora
        ) AS empresas
        OUTER APPLY
        (
            SELECT TOP (1)
                e.IdEmpresa,
                e.CodigoEmpresa,
                e.RazonSocial,
                e.NombreComercial,
                e.Ruc
            FROM dbo.SEG_Empresa AS e
            WHERE e.IdCuentaAdministradora = ca.IdCuentaAdministradora
            ORDER BY
                e.IdEmpresa ASC
        ) AS principal
        LEFT JOIN dbo.SEG_CuentaAdministradoraSuscripcion AS cas
            ON cas.IdCuentaAdministradora = ca.IdCuentaAdministradora
        LEFT JOIN dbo.SEG_UsuarioCuentaAdministradora AS uca
            ON uca.IdCuentaAdministradora = ca.IdCuentaAdministradora
           AND uca.EsCuentaPredeterminada = 1
           AND uca.Estado = 1
        LEFT JOIN dbo.SEG_UsuarioPerfil AS up
            ON up.AspNetUserId = uca.AspNetUserId
        LEFT JOIN dbo.AspNetUsers AS au
            ON au.Id = uca.AspNetUserId
        ORDER BY
            ca.FechaRegistro DESC,
            ca.NombreCuenta ASC;

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
