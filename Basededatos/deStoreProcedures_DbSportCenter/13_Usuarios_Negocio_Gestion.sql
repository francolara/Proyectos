-- =============================================
-- Author:        FRANCO LARA
-- Create date:   27/03/2026
-- Description:   Gestion de usuarios internos por negocio y permisos por modulo.
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
        IF NOT EXISTS (SELECT 1 FROM dbo.ModulosSistema WHERE Codigo = N'CLIENTES')
            INSERT INTO dbo.ModulosSistema (Codigo, Nombre, Activo) VALUES (N'CLIENTES', N'Clientes', 1);
        IF NOT EXISTS (SELECT 1 FROM dbo.ModulosSistema WHERE Codigo = N'ESPACIOS')
            INSERT INTO dbo.ModulosSistema (Codigo, Nombre, Activo) VALUES (N'ESPACIOS', N'Espacios deportivos', 1);
        IF NOT EXISTS (SELECT 1 FROM dbo.ModulosSistema WHERE Codigo = N'RESERVAS')
            INSERT INTO dbo.ModulosSistema (Codigo, Nombre, Activo) VALUES (N'RESERVAS', N'Reservas', 1);
        IF NOT EXISTS (SELECT 1 FROM dbo.ModulosSistema WHERE Codigo = N'SOLICITUDES')
            INSERT INTO dbo.ModulosSistema (Codigo, Nombre, Activo) VALUES (N'SOLICITUDES', N'Solicitudes publicas', 1);
        IF NOT EXISTS (SELECT 1 FROM dbo.ModulosSistema WHERE Codigo = N'PAGOS')
            INSERT INTO dbo.ModulosSistema (Codigo, Nombre, Activo) VALUES (N'PAGOS', N'Pagos', 1);
        IF NOT EXISTS (SELECT 1 FROM dbo.ModulosSistema WHERE Codigo = N'COMPROBANTES')
            INSERT INTO dbo.ModulosSistema (Codigo, Nombre, Activo) VALUES (N'COMPROBANTES', N'Comprobantes electronicos', 1);
        IF NOT EXISTS (SELECT 1 FROM dbo.ModulosSistema WHERE Codigo = N'REPORTES')
            INSERT INTO dbo.ModulosSistema (Codigo, Nombre, Activo) VALUES (N'REPORTES', N'Reportes', 1);
        IF NOT EXISTS (SELECT 1 FROM dbo.ModulosSistema WHERE Codigo = N'USUARIOS')
            INSERT INTO dbo.ModulosSistema (Codigo, Nombre, Activo) VALUES (N'USUARIOS', N'Usuarios del negocio', 1);

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
                      WHEN r.RolNegocio = 2 AND m.Codigo IN (N'DASHBOARD', N'RESERVAS', N'SOLICITUDES', N'PAGOS', N'CLIENTES', N'REPORTES') THEN 1
                      WHEN r.RolNegocio = 3 AND m.Codigo IN (N'DASHBOARD', N'RESERVAS', N'SOLICITUDES', N'CLIENTES') THEN 1
                      WHEN r.RolNegocio = 4 AND m.Codigo IN (N'DASHBOARD', N'PAGOS', N'COMPROBANTES', N'REPORTES') THEN 1
                      WHEN r.RolNegocio = 5 AND m.Codigo IN (N'DASHBOARD', N'RESERVAS', N'ESPACIOS', N'REPORTES') THEN 1
                      ELSE 0 END AS BIT),
            CAST(CASE WHEN r.RolNegocio = 1 THEN 1
                      WHEN r.RolNegocio = 2 AND m.Codigo IN (N'RESERVAS', N'SOLICITUDES', N'PAGOS', N'CLIENTES') THEN 1
                      WHEN r.RolNegocio = 3 AND m.Codigo IN (N'RESERVAS', N'SOLICITUDES', N'CLIENTES') THEN 1
                      WHEN r.RolNegocio = 4 AND m.Codigo IN (N'PAGOS', N'COMPROBANTES') THEN 1
                      ELSE 0 END AS BIT),
            CAST(CASE WHEN r.RolNegocio = 1 THEN 1
                      WHEN r.RolNegocio = 2 AND m.Codigo IN (N'RESERVAS', N'SOLICITUDES', N'PAGOS', N'CLIENTES') THEN 1
                      WHEN r.RolNegocio = 3 AND m.Codigo IN (N'RESERVAS', N'SOLICITUDES', N'CLIENTES') THEN 1
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

CREATE OR ALTER PROCEDURE dbo.Sp_UsuariosNegocio_Listar
    @NegocioId INT
AS
BEGIN
    SET NOCOUNT ON;
    BEGIN TRY
        SELECT
            un.Id AS UsuarioNegocioId,
            un.UsuarioId,
            COALESCE(u.Nombres, N'') AS Nombres,
            COALESCE(u.Apellidos, N'') AS Apellidos,
            COALESCE(u.Email, N'') AS Correo,
            un.RolNegocio,
            un.Activo
        FROM dbo.UsuariosNegocio un
        INNER JOIN dbo.AspNetUsers u ON u.Id = un.UsuarioId
        WHERE un.NegocioId = @NegocioId
        ORDER BY un.Activo DESC, u.Email;
    END TRY
    BEGIN CATCH
        DECLARE @ErrorMessage NVARCHAR(4000), @ErrorSeverity INT, @ErrorState INT;
        SELECT @ErrorMessage = ERROR_MESSAGE(), @ErrorSeverity = ERROR_SEVERITY(), @ErrorState = ERROR_STATE();
        RAISERROR (@ErrorMessage, @ErrorSeverity, @ErrorState);
    END CATCH
END
GO

CREATE OR ALTER PROCEDURE dbo.Sp_UsuariosNegocio_AsignarPorCorreo
    @NegocioId INT,
    @Correo NVARCHAR(256),
    @RolNegocio INT,
    @Usuario NVARCHAR(200)
AS
BEGIN
    SET NOCOUNT ON;
    BEGIN TRY
        DECLARE @UsuarioId NVARCHAR(450);
        SELECT TOP (1) @UsuarioId = u.Id
        FROM dbo.AspNetUsers u
        WHERE u.NormalizedEmail = UPPER(@Correo);

        IF @UsuarioId IS NULL
            RAISERROR('No existe usuario con ese correo en el sistema.', 16, 1);

        IF EXISTS (SELECT 1 FROM dbo.UsuariosNegocio WHERE NegocioId = @NegocioId AND UsuarioId = @UsuarioId)
        BEGIN
            UPDATE dbo.UsuariosNegocio
            SET RolNegocio = @RolNegocio,
                Activo = 1
            WHERE NegocioId = @NegocioId
              AND UsuarioId = @UsuarioId;
        END
        ELSE
        BEGIN
            INSERT INTO dbo.UsuariosNegocio (UsuarioId, NegocioId, RolNegocio, Activo)
            VALUES (@UsuarioId, @NegocioId, @RolNegocio, 1);
        END;

        DECLARE @UsuarioNegocioId INT;
        SELECT TOP (1) @UsuarioNegocioId = Id FROM dbo.UsuariosNegocio WHERE NegocioId = @NegocioId AND UsuarioId = @UsuarioId;

        DECLARE @EntidadIdAudit NVARCHAR(80);
        SET @EntidadIdAudit = CONVERT(NVARCHAR(80), @UsuarioNegocioId);
        EXEC dbo.Sp_Auditoria_Registrar @NegocioId = @NegocioId, @Modulo = N'USUARIOS', @Accion = N'CREATE', @Entidad = N'UsuarioNegocio', @EntidadId = @EntidadIdAudit, @Usuario = @Usuario, @DetalleJson = NULL;
    END TRY
    BEGIN CATCH
        DECLARE @ErrorMessage NVARCHAR(4000), @ErrorSeverity INT, @ErrorState INT;
        SELECT @ErrorMessage = ERROR_MESSAGE(), @ErrorSeverity = ERROR_SEVERITY(), @ErrorState = ERROR_STATE();
        RAISERROR (@ErrorMessage, @ErrorSeverity, @ErrorState);
    END CATCH
END
GO

CREATE OR ALTER PROCEDURE dbo.Sp_UsuariosNegocio_ActualizarRol
    @NegocioId INT,
    @UsuarioNegocioId INT,
    @RolNegocio INT,
    @Usuario NVARCHAR(200)
AS
BEGIN
    SET NOCOUNT ON;
    BEGIN TRY
        UPDATE dbo.UsuariosNegocio
        SET RolNegocio = @RolNegocio
        WHERE Id = @UsuarioNegocioId
          AND NegocioId = @NegocioId;

        IF @@ROWCOUNT > 0
        BEGIN
            DECLARE @EntidadIdAudit NVARCHAR(80);
            SET @EntidadIdAudit = CONVERT(NVARCHAR(80), @UsuarioNegocioId);
            EXEC dbo.Sp_Auditoria_Registrar @NegocioId = @NegocioId, @Modulo = N'USUARIOS', @Accion = N'EDIT', @Entidad = N'UsuarioNegocio', @EntidadId = @EntidadIdAudit, @Usuario = @Usuario, @DetalleJson = NULL;
        END;
    END TRY
    BEGIN CATCH
        DECLARE @ErrorMessage NVARCHAR(4000), @ErrorSeverity INT, @ErrorState INT;
        SELECT @ErrorMessage = ERROR_MESSAGE(), @ErrorSeverity = ERROR_SEVERITY(), @ErrorState = ERROR_STATE();
        RAISERROR (@ErrorMessage, @ErrorSeverity, @ErrorState);
    END CATCH
END
GO

CREATE OR ALTER PROCEDURE dbo.Sp_UsuariosNegocio_Desactivar
    @NegocioId INT,
    @UsuarioNegocioId INT,
    @Usuario NVARCHAR(200)
AS
BEGIN
    SET NOCOUNT ON;
    BEGIN TRY
        UPDATE dbo.UsuariosNegocio
        SET Activo = 0
        WHERE Id = @UsuarioNegocioId
          AND NegocioId = @NegocioId;

        IF @@ROWCOUNT > 0
        BEGIN
            DECLARE @EntidadIdAudit NVARCHAR(80);
            SET @EntidadIdAudit = CONVERT(NVARCHAR(80), @UsuarioNegocioId);
            EXEC dbo.Sp_Auditoria_Registrar @NegocioId = @NegocioId, @Modulo = N'USUARIOS', @Accion = N'DELETE', @Entidad = N'UsuarioNegocio', @EntidadId = @EntidadIdAudit, @Usuario = @Usuario, @DetalleJson = NULL;
        END;
    END TRY
    BEGIN CATCH
        DECLARE @ErrorMessage NVARCHAR(4000), @ErrorSeverity INT, @ErrorState INT;
        SELECT @ErrorMessage = ERROR_MESSAGE(), @ErrorSeverity = ERROR_SEVERITY(), @ErrorState = ERROR_STATE();
        RAISERROR (@ErrorMessage, @ErrorSeverity, @ErrorState);
    END CATCH
END
GO

CREATE OR ALTER PROCEDURE dbo.Sp_UsuariosNegocio_PermisosListar
    @NegocioId INT,
    @UsuarioNegocioId INT
AS
BEGIN
    SET NOCOUNT ON;
    BEGIN TRY
        IF NOT EXISTS (SELECT 1 FROM dbo.UsuariosNegocio WHERE Id = @UsuarioNegocioId AND NegocioId = @NegocioId)
            RAISERROR('UsuarioNegocio invalido para el negocio.', 16, 1);

        DECLARE @RolNegocio INT;
        SELECT @RolNegocio = RolNegocio FROM dbo.UsuariosNegocio WHERE Id = @UsuarioNegocioId;

        SELECT
            m.Id AS ModuloSistemaId,
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

CREATE OR ALTER PROCEDURE dbo.Sp_UsuariosNegocio_PermisoGuardar
    @NegocioId INT,
    @UsuarioNegocioId INT,
    @ModuloSistemaId INT,
    @PuedeVer BIT,
    @PuedeCrear BIT,
    @PuedeEditar BIT,
    @PuedeEliminar BIT,
    @Usuario NVARCHAR(200)
AS
BEGIN
    SET NOCOUNT ON;
    BEGIN TRY
        IF NOT EXISTS (SELECT 1 FROM dbo.UsuariosNegocio WHERE Id = @UsuarioNegocioId AND NegocioId = @NegocioId)
            RAISERROR('UsuarioNegocio invalido para el negocio.', 16, 1);

        IF EXISTS (SELECT 1 FROM dbo.UsuariosNegocioPermiso WHERE UsuarioNegocioId = @UsuarioNegocioId AND ModuloSistemaId = @ModuloSistemaId)
        BEGIN
            UPDATE dbo.UsuariosNegocioPermiso
            SET PuedeVer = @PuedeVer,
                PuedeCrear = @PuedeCrear,
                PuedeEditar = @PuedeEditar,
                PuedeEliminar = @PuedeEliminar
            WHERE UsuarioNegocioId = @UsuarioNegocioId
              AND ModuloSistemaId = @ModuloSistemaId;
        END
        ELSE
        BEGIN
            INSERT INTO dbo.UsuariosNegocioPermiso
            (
                UsuarioNegocioId, ModuloSistemaId, PuedeVer, PuedeCrear, PuedeEditar, PuedeEliminar
            )
            VALUES
            (
                @UsuarioNegocioId, @ModuloSistemaId, @PuedeVer, @PuedeCrear, @PuedeEditar, @PuedeEliminar
            );
        END;

        DECLARE @EntidadIdAudit NVARCHAR(80);
        SET @EntidadIdAudit = CONCAT(CONVERT(NVARCHAR(30), @UsuarioNegocioId), N'-', CONVERT(NVARCHAR(30), @ModuloSistemaId));
        EXEC dbo.Sp_Auditoria_Registrar @NegocioId = @NegocioId, @Modulo = N'USUARIOS', @Accion = N'EDIT', @Entidad = N'UsuarioNegocioPermiso', @EntidadId = @EntidadIdAudit, @Usuario = @Usuario, @DetalleJson = NULL;
    END TRY
    BEGIN CATCH
        DECLARE @ErrorMessage NVARCHAR(4000), @ErrorSeverity INT, @ErrorState INT;
        SELECT @ErrorMessage = ERROR_MESSAGE(), @ErrorSeverity = ERROR_SEVERITY(), @ErrorState = ERROR_STATE();
        RAISERROR (@ErrorMessage, @ErrorSeverity, @ErrorState);
    END CATCH
END
GO
