-- =============================================
-- Author:        FRANCO LARA / Codex
-- Create date:   10/07/2026
-- Description:   Resuelve el contexto de acceso inicial del usuario autenticado segun superadmin, cuenta administradora y empresas asignadas.
-- =============================================

CREATE OR ALTER PROCEDURE dbo.usp_SEG_ObtenerContextoLoginUsuario
    @AspNetUserId NVARCHAR(450)
AS
BEGIN

    SET NOCOUNT ON;

    BEGIN TRY

        DECLARE @EsSuperAdmin BIT = 0;

        IF EXISTS
        (
            SELECT 1
            FROM dbo.AspNetUserRoles AS ur
            INNER JOIN dbo.AspNetRoles AS r
                ON r.Id = ur.RoleId
            WHERE ur.UserId = @AspNetUserId
              AND r.Name = N'SuperAdmin'
        )
        BEGIN
            SET @EsSuperAdmin = 1;
        END;

        IF @EsSuperAdmin = 1
        BEGIN
            SELECT
                CAST(1 AS BIT) AS TieneAcceso,
                CAST(1 AS BIT) AS EsSuperAdmin,
                CAST(NULL AS INT) AS IdCuentaAdministradora,
                CAST(NULL AS VARCHAR(20)) AS CodigoCuenta,
                CAST(NULL AS NVARCHAR(200)) AS NombreCuenta,
                CAST(N'SUPERADMIN' AS NVARCHAR(30)) AS RolCuenta,
                CAST(0 AS INT) AS CantidadEmpresasAsignadas,
                CAST(NULL AS INT) AS IdEmpresaPredeterminada,
                CAST(NULL AS NVARCHAR(200)) AS RazonSocialEmpresaPredeterminada,
                CAST(0 AS BIT) AS DebeSeleccionarEmpresa,
                CAST(0 AS BIT) AS SoloModulosCuenta,
                CAST(N'Usuario con acceso de plataforma.' AS NVARCHAR(250)) AS Mensaje;

            RETURN;
        END;

        DECLARE
            @IdCuentaAdministradora INT,
            @CodigoCuenta VARCHAR(20),
            @NombreCuenta NVARCHAR(200),
            @RolCuenta NVARCHAR(30);

        SELECT TOP (1)
            @IdCuentaAdministradora = uca.IdCuentaAdministradora,
            @CodigoCuenta = ca.CodigoCuenta,
            @NombreCuenta = ca.NombreCuenta,
            @RolCuenta = uca.RolCuenta
        FROM dbo.SEG_UsuarioCuentaAdministradora AS uca
        INNER JOIN dbo.SEG_CuentaAdministradora AS ca
            ON ca.IdCuentaAdministradora = uca.IdCuentaAdministradora
        WHERE uca.AspNetUserId = @AspNetUserId
          AND uca.Estado = 1
          AND ca.Estado = 1
        ORDER BY
            uca.EsCuentaPredeterminada DESC,
            uca.IdUsuarioCuentaAdministradora ASC;

        IF @IdCuentaAdministradora IS NULL
        BEGIN
            SELECT
                CAST(0 AS BIT) AS TieneAcceso,
                CAST(0 AS BIT) AS EsSuperAdmin,
                CAST(NULL AS INT) AS IdCuentaAdministradora,
                CAST(NULL AS VARCHAR(20)) AS CodigoCuenta,
                CAST(NULL AS NVARCHAR(200)) AS NombreCuenta,
                CAST(NULL AS NVARCHAR(30)) AS RolCuenta,
                CAST(0 AS INT) AS CantidadEmpresasAsignadas,
                CAST(NULL AS INT) AS IdEmpresaPredeterminada,
                CAST(NULL AS NVARCHAR(200)) AS RazonSocialEmpresaPredeterminada,
                CAST(0 AS BIT) AS DebeSeleccionarEmpresa,
                CAST(0 AS BIT) AS SoloModulosCuenta,
                CAST(N'El usuario no esta vinculado a una cuenta administradora activa.' AS NVARCHAR(250)) AS Mensaje;

            RETURN;
        END;

        DECLARE @Empresas TABLE
        (
            IdEmpresa INT NOT NULL,
            RazonSocial NVARCHAR(200) NOT NULL,
            EsEmpresaPredeterminada BIT NOT NULL
        );

        INSERT INTO @Empresas
        (
            IdEmpresa,
            RazonSocial,
            EsEmpresaPredeterminada
        )
        SELECT
            ue.IdEmpresa,
            e.RazonSocial,
            ue.EsEmpresaPredeterminada
        FROM dbo.SEG_UsuarioEmpresa AS ue
        INNER JOIN dbo.SEG_Empresa AS e
            ON e.IdEmpresa = ue.IdEmpresa
        WHERE ue.AspNetUserId = @AspNetUserId
          AND ue.Estado = 1
          AND e.Estado = 1
          AND e.IdCuentaAdministradora = @IdCuentaAdministradora;

        DECLARE
            @CantidadEmpresasAsignadas INT = (SELECT COUNT(1) FROM @Empresas),
            @IdEmpresaPredeterminada INT,
            @RazonSocialEmpresaPredeterminada NVARCHAR(200);

        SELECT TOP (1)
            @IdEmpresaPredeterminada = emp.IdEmpresa,
            @RazonSocialEmpresaPredeterminada = emp.RazonSocial
        FROM @Empresas AS emp
        ORDER BY
            emp.EsEmpresaPredeterminada DESC,
            emp.RazonSocial ASC;

        SELECT
            CAST(1 AS BIT) AS TieneAcceso,
            CAST(0 AS BIT) AS EsSuperAdmin,
            @IdCuentaAdministradora AS IdCuentaAdministradora,
            @CodigoCuenta AS CodigoCuenta,
            @NombreCuenta AS NombreCuenta,
            @RolCuenta AS RolCuenta,
            @CantidadEmpresasAsignadas AS CantidadEmpresasAsignadas,
            @IdEmpresaPredeterminada AS IdEmpresaPredeterminada,
            @RazonSocialEmpresaPredeterminada AS RazonSocialEmpresaPredeterminada,
            CAST(CASE WHEN @CantidadEmpresasAsignadas > 1 THEN 1 ELSE 0 END AS BIT) AS DebeSeleccionarEmpresa,
            CAST(CASE WHEN @CantidadEmpresasAsignadas = 0 THEN 1 ELSE 0 END AS BIT) AS SoloModulosCuenta,
            CAST(
                CASE
                    WHEN @CantidadEmpresasAsignadas = 0 THEN N'El usuario tiene acceso a la cuenta administradora, pero aun no tiene empresas asignadas.'
                    WHEN @CantidadEmpresasAsignadas = 1 THEN N'El usuario tiene una empresa asignada y puede ingresar directamente.'
                    ELSE N'El usuario tiene multiples empresas asignadas y debe seleccionar una antes de operar.'
                END
                AS NVARCHAR(250)
            ) AS Mensaje;

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
