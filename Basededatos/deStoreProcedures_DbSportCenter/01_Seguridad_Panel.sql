-- =============================================
-- Author:        FRANCO LARA
-- Create date:   26/03/2026
-- Description:   Seguridad y panel por negocio (roles, permisos y metricas).
-- =============================================

CREATE OR ALTER PROCEDURE dbo.Sp_Seguridad_SeedModulosPermisosBase
AS
BEGIN
    SET NOCOUNT ON;
    BEGIN TRY
        IF NOT EXISTS (SELECT 1 FROM dbo.ModulosSistema WHERE Codigo = N'DASHBOARD')
            INSERT INTO dbo.ModulosSistema (Codigo, Nombre, Activo) VALUES (N'DASHBOARD', N'Dashboard', 1);
        IF NOT EXISTS (SELECT 1 FROM dbo.ModulosSistema WHERE Codigo = N'SEDES')
            INSERT INTO dbo.ModulosSistema (Codigo, Nombre, Activo) VALUES (N'SEDES', N'Sedes', 1);
        IF NOT EXISTS (SELECT 1 FROM dbo.ModulosSistema WHERE Codigo = N'ESPACIOS')
            INSERT INTO dbo.ModulosSistema (Codigo, Nombre, Activo) VALUES (N'ESPACIOS', N'Espacios deportivos', 1);
        IF NOT EXISTS (SELECT 1 FROM dbo.ModulosSistema WHERE Codigo = N'RESERVAS')
            INSERT INTO dbo.ModulosSistema (Codigo, Nombre, Activo) VALUES (N'RESERVAS', N'Reservas', 1);
        IF NOT EXISTS (SELECT 1 FROM dbo.ModulosSistema WHERE Codigo = N'PAGOS')
            INSERT INTO dbo.ModulosSistema (Codigo, Nombre, Activo) VALUES (N'PAGOS', N'Pagos', 1);
        IF NOT EXISTS (SELECT 1 FROM dbo.ModulosSistema WHERE Codigo = N'COMPROBANTES')
            INSERT INTO dbo.ModulosSistema (Codigo, Nombre, Activo) VALUES (N'COMPROBANTES', N'Comprobantes electronicos', 1);

        ;WITH Roles AS
        (
            SELECT CAST(1 AS INT) AS RolNegocio UNION ALL
            SELECT 2 UNION ALL
            SELECT 3 UNION ALL
            SELECT 4 UNION ALL
            SELECT 5
        )
        INSERT INTO dbo.RolesNegocioPermiso (RolNegocio, ModuloSistemaId, PuedeVer, PuedeCrear, PuedeEditar, PuedeEliminar)
        SELECT
            r.RolNegocio,
            m.Id,
            CAST(CASE WHEN r.RolNegocio = 1 THEN 1
                      WHEN r.RolNegocio = 2 AND m.Codigo IN (N'DASHBOARD', N'RESERVAS', N'PAGOS') THEN 1
                      WHEN r.RolNegocio = 3 AND m.Codigo IN (N'DASHBOARD', N'RESERVAS') THEN 1
                      WHEN r.RolNegocio = 4 AND m.Codigo IN (N'DASHBOARD', N'PAGOS', N'COMPROBANTES') THEN 1
                      WHEN r.RolNegocio = 5 AND m.Codigo IN (N'DASHBOARD', N'RESERVAS', N'ESPACIOS') THEN 1
                      ELSE 0 END AS BIT),
            CAST(CASE WHEN r.RolNegocio = 1 THEN 1
                      WHEN r.RolNegocio = 2 AND m.Codigo IN (N'RESERVAS', N'PAGOS') THEN 1
                      WHEN r.RolNegocio = 3 AND m.Codigo = N'RESERVAS' THEN 1
                      WHEN r.RolNegocio = 4 AND m.Codigo IN (N'PAGOS', N'COMPROBANTES') THEN 1
                      ELSE 0 END AS BIT),
            CAST(CASE WHEN r.RolNegocio = 1 THEN 1
                      WHEN r.RolNegocio = 2 AND m.Codigo IN (N'RESERVAS', N'PAGOS') THEN 1
                      WHEN r.RolNegocio = 3 AND m.Codigo = N'RESERVAS' THEN 1
                      WHEN r.RolNegocio = 4 AND m.Codigo IN (N'PAGOS', N'COMPROBANTES') THEN 1
                      WHEN r.RolNegocio = 5 AND m.Codigo IN (N'RESERVAS', N'ESPACIOS') THEN 1
                      ELSE 0 END AS BIT),
            CAST(CASE WHEN r.RolNegocio = 1 THEN 1 ELSE 0 END AS BIT)
        FROM Roles r
        CROSS JOIN dbo.ModulosSistema m
        WHERE NOT EXISTS (
            SELECT 1
            FROM dbo.RolesNegocioPermiso rp
            WHERE rp.RolNegocio = r.RolNegocio
              AND rp.ModuloSistemaId = m.Id
        );
    END TRY
    BEGIN CATCH
        DECLARE @ErrorMessage NVARCHAR(4000), @ErrorSeverity INT, @ErrorState INT;
        SELECT @ErrorMessage = ERROR_MESSAGE(), @ErrorSeverity = ERROR_SEVERITY(), @ErrorState = ERROR_STATE();
        RAISERROR (@ErrorMessage, @ErrorSeverity, @ErrorState);
    END CATCH
END
GO

CREATE OR ALTER PROCEDURE dbo.Sp_Seguridad_ObtenerContextoModulo
    @UsuarioId NVARCHAR(450),
    @NegocioId INT,
    @ModuloCodigo NVARCHAR(50)
AS
BEGIN
    SET NOCOUNT ON;
    BEGIN TRY
        DECLARE @UsuarioNegocioId INT, @RolNegocio INT, @NegocioNombre NVARCHAR(200), @ModuloId INT, @ModuloNombre NVARCHAR(120);
        DECLARE @PuedeVer BIT = 0, @PuedeCrear BIT = 0, @PuedeEditar BIT = 0, @PuedeEliminar BIT = 0;

        SELECT
            @UsuarioNegocioId = un.Id,
            @RolNegocio = un.RolNegocio,
            @NegocioNombre = n.NombreComercial
        FROM dbo.UsuariosNegocio un
        INNER JOIN dbo.Negocios n ON n.Id = un.NegocioId
        WHERE un.UsuarioId = @UsuarioId
          AND un.NegocioId = @NegocioId
          AND un.Activo = 1
          AND n.Activo = 1;

        IF @UsuarioNegocioId IS NULL
        BEGIN
            SELECT CAST(0 AS BIT), @NegocioId, N'', @ModuloCodigo, N'', N'', CAST(0 AS BIT), CAST(0 AS BIT), CAST(0 AS BIT), CAST(0 AS BIT), N'Usuario sin acceso al negocio';
            RETURN;
        END;

        SELECT @ModuloId = m.Id, @ModuloNombre = m.Nombre
        FROM dbo.ModulosSistema m
        WHERE m.Codigo = @ModuloCodigo AND m.Activo = 1;

        IF @ModuloId IS NULL
        BEGIN
            SELECT CAST(0 AS BIT), @NegocioId, @NegocioNombre, @ModuloCodigo, N'', CAST(@RolNegocio AS NVARCHAR(20)), CAST(0 AS BIT), CAST(0 AS BIT), CAST(0 AS BIT), CAST(0 AS BIT), N'Modulo no configurado';
            RETURN;
        END;

        SELECT
            @PuedeVer = rp.PuedeVer,
            @PuedeCrear = rp.PuedeCrear,
            @PuedeEditar = rp.PuedeEditar,
            @PuedeEliminar = rp.PuedeEliminar
        FROM dbo.RolesNegocioPermiso rp
        WHERE rp.RolNegocio = @RolNegocio
          AND rp.ModuloSistemaId = @ModuloId;

        SELECT
            @PuedeVer = COALESCE(up.PuedeVer, @PuedeVer),
            @PuedeCrear = COALESCE(up.PuedeCrear, @PuedeCrear),
            @PuedeEditar = COALESCE(up.PuedeEditar, @PuedeEditar),
            @PuedeEliminar = COALESCE(up.PuedeEliminar, @PuedeEliminar)
        FROM dbo.UsuariosNegocioPermiso up
        WHERE up.UsuarioNegocioId = @UsuarioNegocioId
          AND up.ModuloSistemaId = @ModuloId;

        SELECT
            CAST(CASE WHEN @PuedeVer = 1 THEN 1 ELSE 0 END AS BIT) AS Autorizado,
            @NegocioId,
            @NegocioNombre,
            @ModuloCodigo,
            @ModuloNombre,
            CAST(@RolNegocio AS NVARCHAR(20)) AS RolActual,
            @PuedeVer,
            @PuedeCrear,
            @PuedeEditar,
            @PuedeEliminar,
            CAST(NULL AS NVARCHAR(200)) AS Mensaje;
    END TRY
    BEGIN CATCH
        DECLARE @ErrorMessage NVARCHAR(4000), @ErrorSeverity INT, @ErrorState INT;
        SELECT @ErrorMessage = ERROR_MESSAGE(), @ErrorSeverity = ERROR_SEVERITY(), @ErrorState = ERROR_STATE();
        RAISERROR (@ErrorMessage, @ErrorSeverity, @ErrorState);
    END CATCH
END
GO

CREATE OR ALTER PROCEDURE dbo.Sp_Panel_ListarNegociosUsuario
    @UsuarioId NVARCHAR(450)
AS
BEGIN
    SET NOCOUNT ON;
    BEGIN TRY
        SELECT
            un.NegocioId,
            n.NombreComercial,
            CAST(un.RolNegocio AS NVARCHAR(20)) AS Rol
        FROM dbo.UsuariosNegocio un
        INNER JOIN dbo.Negocios n ON n.Id = un.NegocioId
        WHERE un.UsuarioId = @UsuarioId
          AND un.Activo = 1
          AND n.Activo = 1
        ORDER BY n.NombreComercial;
    END TRY
    BEGIN CATCH
        DECLARE @ErrorMessage NVARCHAR(4000), @ErrorSeverity INT, @ErrorState INT;
        SELECT @ErrorMessage = ERROR_MESSAGE(), @ErrorSeverity = ERROR_SEVERITY(), @ErrorState = ERROR_STATE();
        RAISERROR (@ErrorMessage, @ErrorSeverity, @ErrorState);
    END CATCH
END
GO

CREATE OR ALTER PROCEDURE dbo.Sp_Panel_ObtenerRolUsuario
    @UsuarioId NVARCHAR(450),
    @NegocioId INT
AS
BEGIN
    SET NOCOUNT ON;
    BEGIN TRY
        SELECT TOP (1)
            CAST(un.RolNegocio AS NVARCHAR(20))
        FROM dbo.UsuariosNegocio un
        WHERE un.UsuarioId = @UsuarioId
          AND un.NegocioId = @NegocioId
          AND un.Activo = 1;
    END TRY
    BEGIN CATCH
        DECLARE @ErrorMessage NVARCHAR(4000), @ErrorSeverity INT, @ErrorState INT;
        SELECT @ErrorMessage = ERROR_MESSAGE(), @ErrorSeverity = ERROR_SEVERITY(), @ErrorState = ERROR_STATE();
        RAISERROR (@ErrorMessage, @ErrorSeverity, @ErrorState);
    END CATCH
END
GO

CREATE OR ALTER PROCEDURE dbo.Sp_Panel_ListarModulosPermitidos
    @UsuarioId NVARCHAR(450),
    @NegocioId INT
AS
BEGIN
    SET NOCOUNT ON;
    BEGIN TRY
        DECLARE @UsuarioNegocioId INT, @RolNegocio INT;

        SELECT @UsuarioNegocioId = un.Id, @RolNegocio = un.RolNegocio
        FROM dbo.UsuariosNegocio un
        WHERE un.UsuarioId = @UsuarioId
          AND un.NegocioId = @NegocioId
          AND un.Activo = 1;

        SELECT
            m.Id,
            m.Codigo,
            m.Nombre,
            COALESCE(up.PuedeVer, rp.PuedeVer) AS PuedeVer,
            COALESCE(up.PuedeCrear, rp.PuedeCrear) AS PuedeCrear,
            COALESCE(up.PuedeEditar, rp.PuedeEditar) AS PuedeEditar,
            COALESCE(up.PuedeEliminar, rp.PuedeEliminar) AS PuedeEliminar
        FROM dbo.ModulosSistema m
        INNER JOIN dbo.RolesNegocioPermiso rp ON rp.ModuloSistemaId = m.Id AND rp.RolNegocio = @RolNegocio
        LEFT JOIN dbo.UsuariosNegocioPermiso up ON up.ModuloSistemaId = m.Id AND up.UsuarioNegocioId = @UsuarioNegocioId
        WHERE m.Activo = 1
        ORDER BY m.Nombre;
    END TRY
    BEGIN CATCH
        DECLARE @ErrorMessage NVARCHAR(4000), @ErrorSeverity INT, @ErrorState INT;
        SELECT @ErrorMessage = ERROR_MESSAGE(), @ErrorSeverity = ERROR_SEVERITY(), @ErrorState = ERROR_STATE();
        RAISERROR (@ErrorMessage, @ErrorSeverity, @ErrorState);
    END CATCH
END
GO

CREATE OR ALTER PROCEDURE dbo.Sp_Panel_ObtenerMetricas
    @NegocioId INT,
    @Fecha DATE
AS
BEGIN
    SET NOCOUNT ON;
    BEGIN TRY
        SELECT
            (SELECT COUNT(1) FROM dbo.Sedes s WHERE s.NegocioId = @NegocioId) AS TotalSedes,
            (SELECT COUNT(1)
             FROM dbo.EspaciosDeportivos e
             INNER JOIN dbo.Sedes s ON s.Id = e.SedeId
             WHERE s.NegocioId = @NegocioId) AS TotalEspacios,
            (SELECT COUNT(1)
             FROM dbo.Reservas r
             INNER JOIN dbo.EspaciosDeportivos e ON e.Id = r.EspacioDeportivoId
             INNER JOIN dbo.Sedes s ON s.Id = e.SedeId
             WHERE s.NegocioId = @NegocioId
               AND r.Fecha = @Fecha) AS ReservasHoy,
            (SELECT COALESCE(SUM(p.Monto), 0)
             FROM dbo.Pagos p
             INNER JOIN dbo.Reservas r ON r.Id = p.ReservaId
             INNER JOIN dbo.EspaciosDeportivos e ON e.Id = r.EspacioDeportivoId
             INNER JOIN dbo.Sedes s ON s.Id = e.SedeId
             WHERE s.NegocioId = @NegocioId
               AND r.Fecha = @Fecha) AS IngresosHoy;
    END TRY
    BEGIN CATCH
        DECLARE @ErrorMessage NVARCHAR(4000), @ErrorSeverity INT, @ErrorState INT;
        SELECT @ErrorMessage = ERROR_MESSAGE(), @ErrorSeverity = ERROR_SEVERITY(), @ErrorState = ERROR_STATE();
        RAISERROR (@ErrorMessage, @ErrorSeverity, @ErrorState);
    END CATCH
END
GO