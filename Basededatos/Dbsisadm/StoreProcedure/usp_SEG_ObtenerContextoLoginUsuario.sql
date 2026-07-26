-- =============================================
-- Author:        FRANCO LARA / Codex
-- Create date:   10/07/2026
-- Description:   Resuelve el contexto de acceso inicial del usuario autenticado segun superadmin, cuenta administradora y empresas asignadas.
-- =============================================
-- =============================================
-- Author:        FRANCO LARA / Codex
-- Create date:   25/07/2026
-- Description:   Incorpora la vigencia comercial al contexto de login para validar centralmente prueba, plan, gracia, suspension y baja sin modificar los codigos almacenados.
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
                CAST(NULL AS NVARCHAR(256)) AS CorreoPrincipal,
                CAST(NULL AS NVARCHAR(30)) AS TelefonoPrincipal,
                CAST(NULL AS BIT) AS EstadoCuenta,
                CAST(N'SUPERADMIN' AS NVARCHAR(30)) AS RolCuenta,
                CAST(0 AS INT) AS CantidadEmpresasAsignadas,
                CAST(NULL AS INT) AS IdEmpresaPredeterminada,
                CAST(NULL AS NVARCHAR(200)) AS RazonSocialEmpresaPredeterminada,
                CAST(0 AS BIT) AS DebeSeleccionarEmpresa,
                CAST(0 AS BIT) AS SoloModulosCuenta,
                CAST(NULL AS INT) AS IdCuentaAdministradoraSuscripcion,
                CAST(NULL AS NVARCHAR(50)) AS TipoPlan,
                CAST(NULL AS NVARCHAR(20)) AS EstadoSuscripcion,
                CAST(NULL AS BIT) AS EsPrueba,
                CAST(NULL AS DATE) AS FechaInicioPrueba,
                CAST(NULL AS DATE) AS FechaFinPrueba,
                CAST(NULL AS DATE) AS FechaInicioPlan,
                CAST(NULL AS DATE) AS FechaFinPlan,
                CAST(NULL AS INT) AS DiasGracia,
                CAST(NULL AS DATE) AS FechaFinGracia,
                CAST(NULL AS INT) AS EmpresasPermitidas,
                CAST(NULL AS INT) AS UsuariosPermitidos,
                CAST(NULL AS BIT) AS ActivoSuscripcion,
                CAST(NULL AS NVARCHAR(500)) AS ObservacionSuscripcion,
                CAST(N'Usuario con acceso de plataforma.' AS NVARCHAR(250)) AS Mensaje;

            RETURN;
        END;

        DECLARE
            @IdCuentaAdministradora INT,
            @CodigoCuenta VARCHAR(20),
            @NombreCuenta NVARCHAR(200),
            @CorreoPrincipal NVARCHAR(256),
            @TelefonoPrincipal NVARCHAR(30),
            @EstadoCuenta BIT,
            @RolCuenta NVARCHAR(30);

        SELECT TOP (1)
            @IdCuentaAdministradora = uca.IdCuentaAdministradora,
            @CodigoCuenta = ca.CodigoCuenta,
            @NombreCuenta = ca.NombreCuenta,
            @CorreoPrincipal = ca.CorreoPrincipal,
            @TelefonoPrincipal = ca.TelefonoPrincipal,
            @EstadoCuenta = ca.Estado,
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
                CAST(NULL AS NVARCHAR(256)) AS CorreoPrincipal,
                CAST(NULL AS NVARCHAR(30)) AS TelefonoPrincipal,
                CAST(NULL AS BIT) AS EstadoCuenta,
                CAST(NULL AS NVARCHAR(30)) AS RolCuenta,
                CAST(0 AS INT) AS CantidadEmpresasAsignadas,
                CAST(NULL AS INT) AS IdEmpresaPredeterminada,
                CAST(NULL AS NVARCHAR(200)) AS RazonSocialEmpresaPredeterminada,
                CAST(0 AS BIT) AS DebeSeleccionarEmpresa,
                CAST(0 AS BIT) AS SoloModulosCuenta,
                CAST(NULL AS INT) AS IdCuentaAdministradoraSuscripcion,
                CAST(NULL AS NVARCHAR(50)) AS TipoPlan,
                CAST(NULL AS NVARCHAR(20)) AS EstadoSuscripcion,
                CAST(NULL AS BIT) AS EsPrueba,
                CAST(NULL AS DATE) AS FechaInicioPrueba,
                CAST(NULL AS DATE) AS FechaFinPrueba,
                CAST(NULL AS DATE) AS FechaInicioPlan,
                CAST(NULL AS DATE) AS FechaFinPlan,
                CAST(NULL AS INT) AS DiasGracia,
                CAST(NULL AS DATE) AS FechaFinGracia,
                CAST(NULL AS INT) AS EmpresasPermitidas,
                CAST(NULL AS INT) AS UsuariosPermitidos,
                CAST(NULL AS BIT) AS ActivoSuscripcion,
                CAST(NULL AS NVARCHAR(500)) AS ObservacionSuscripcion,
                CAST(N'El usuario no esta vinculado a una cuenta administradora activa.' AS NVARCHAR(250)) AS Mensaje;

            RETURN;
        END;

        DECLARE
            @IdCuentaAdministradoraSuscripcion INT,
            @TipoPlan NVARCHAR(50),
            @EstadoSuscripcion NVARCHAR(20),
            @EsPrueba BIT,
            @FechaInicioPrueba DATE,
            @FechaFinPrueba DATE,
            @FechaInicioPlan DATE,
            @FechaFinPlan DATE,
            @DiasGracia INT,
            @FechaFinGracia DATE,
            @EmpresasPermitidas INT,
            @UsuariosPermitidos INT,
            @ActivoSuscripcion BIT,
            @ObservacionSuscripcion NVARCHAR(500);

        SELECT
            @IdCuentaAdministradoraSuscripcion = cas.IdCuentaAdministradoraSuscripcion,
            @TipoPlan = cas.TipoPlan,
            @EstadoSuscripcion = cas.EstadoSuscripcion,
            @EsPrueba = cas.EsPrueba,
            @FechaInicioPrueba = cas.FechaInicioPrueba,
            @FechaFinPrueba = cas.FechaFinPrueba,
            @FechaInicioPlan = cas.FechaInicioPlan,
            @FechaFinPlan = cas.FechaFinPlan,
            @DiasGracia = cas.DiasGracia,
            @FechaFinGracia = cas.FechaFinGracia,
            @EmpresasPermitidas = cas.EmpresasPermitidas,
            @UsuariosPermitidos = cas.UsuariosPermitidos,
            @ActivoSuscripcion = cas.Activo,
            @ObservacionSuscripcion = cas.Observacion
        FROM dbo.SEG_CuentaAdministradoraSuscripcion AS cas
        WHERE cas.IdCuentaAdministradora = @IdCuentaAdministradora;

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
            @CorreoPrincipal AS CorreoPrincipal,
            @TelefonoPrincipal AS TelefonoPrincipal,
            @EstadoCuenta AS EstadoCuenta,
            @RolCuenta AS RolCuenta,
            @CantidadEmpresasAsignadas AS CantidadEmpresasAsignadas,
            @IdEmpresaPredeterminada AS IdEmpresaPredeterminada,
            @RazonSocialEmpresaPredeterminada AS RazonSocialEmpresaPredeterminada,
            CAST(CASE WHEN @CantidadEmpresasAsignadas > 1 THEN 1 ELSE 0 END AS BIT) AS DebeSeleccionarEmpresa,
            CAST(CASE WHEN @CantidadEmpresasAsignadas = 0 THEN 1 ELSE 0 END AS BIT) AS SoloModulosCuenta,
            @IdCuentaAdministradoraSuscripcion AS IdCuentaAdministradoraSuscripcion,
            @TipoPlan AS TipoPlan,
            @EstadoSuscripcion AS EstadoSuscripcion,
            @EsPrueba AS EsPrueba,
            @FechaInicioPrueba AS FechaInicioPrueba,
            @FechaFinPrueba AS FechaFinPrueba,
            @FechaInicioPlan AS FechaInicioPlan,
            @FechaFinPlan AS FechaFinPlan,
            @DiasGracia AS DiasGracia,
            @FechaFinGracia AS FechaFinGracia,
            @EmpresasPermitidas AS EmpresasPermitidas,
            @UsuariosPermitidos AS UsuariosPermitidos,
            @ActivoSuscripcion AS ActivoSuscripcion,
            @ObservacionSuscripcion AS ObservacionSuscripcion,
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
