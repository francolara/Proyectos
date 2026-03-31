-- =============================================
-- Author:        FRANCO LARA
-- Create date:   27/03/2026
-- Description:   Modulo clientes por negocio (seguridad + CRUD + combos filtrados) y columnas de auditoria.
-- =============================================
-- =============================================
-- Author:        FRANCO LARA
-- Create date:   30/03/2026
-- Description:   Ajusta update/delete de clientes para devolver error controlado cuando no existe registro para el negocio y valida duplicado por numero de documento dentro del negocio.
-- Firma:         Codex - 30/03/2026 | Centraliza validacion de existencia y duplicidad (numero de documento) en SP para crear/actualizar cliente.
-- =============================================

IF COL_LENGTH('dbo.Clientes', 'DireccionFiscal') IS NULL
BEGIN
    ALTER TABLE dbo.Clientes ADD DireccionFiscal NVARCHAR(250) NULL;
END;
GO

IF COL_LENGTH('dbo.Clientes', 'FechaCreacion') IS NULL
BEGIN
    ALTER TABLE dbo.Clientes ADD FechaCreacion DATETIME2 NOT NULL CONSTRAINT DF_Clientes_FechaCreacion DEFAULT (SYSUTCDATETIME());
END;
GO

IF COL_LENGTH('dbo.Clientes', 'UsuarioCreacion') IS NULL
BEGIN
    ALTER TABLE dbo.Clientes ADD UsuarioCreacion NVARCHAR(200) NULL;
END;
GO

IF COL_LENGTH('dbo.Clientes', 'FechaActualizacion') IS NULL
BEGIN
    ALTER TABLE dbo.Clientes ADD FechaActualizacion DATETIME2 NULL;
END;
GO

IF COL_LENGTH('dbo.Clientes', 'UsuarioActualizacion') IS NULL
BEGIN
    ALTER TABLE dbo.Clientes ADD UsuarioActualizacion NVARCHAR(200) NULL;
END;
GO

IF OBJECT_ID(N'dbo.NegocioClientes', N'U') IS NULL
BEGIN
    CREATE TABLE dbo.NegocioClientes
    (
        NegocioId INT NOT NULL,
        ClienteId INT NOT NULL,
        Activo BIT NOT NULL CONSTRAINT DF_NegocioClientes_Activo DEFAULT (1),
        FechaRegistro DATETIME2 NOT NULL CONSTRAINT DF_NegocioClientes_FechaRegistro DEFAULT (SYSUTCDATETIME()),
        UsuarioCreacion NVARCHAR(200) NULL,
        CONSTRAINT PK_NegocioClientes PRIMARY KEY (NegocioId, ClienteId),
        CONSTRAINT FK_NegocioClientes_Negocios_NegocioId FOREIGN KEY (NegocioId) REFERENCES dbo.Negocios (Id),
        CONSTRAINT FK_NegocioClientes_Clientes_ClienteId FOREIGN KEY (ClienteId) REFERENCES dbo.Clientes (Id)
    );
END;
GO

IF NOT EXISTS (SELECT 1 FROM sys.indexes WHERE object_id = OBJECT_ID(N'dbo.NegocioClientes') AND name = N'IX_NegocioClientes_ClienteId')
BEGIN
    CREATE INDEX IX_NegocioClientes_ClienteId ON dbo.NegocioClientes (ClienteId);
END;
GO

INSERT INTO dbo.NegocioClientes (NegocioId, ClienteId, Activo, FechaRegistro, UsuarioCreacion)
SELECT DISTINCT
    s.NegocioId,
    r.ClienteId,
    1,
    SYSUTCDATETIME(),
    N'sistema'
FROM dbo.Reservas r
INNER JOIN dbo.EspaciosDeportivos e ON e.Id = r.EspacioDeportivoId
INNER JOIN dbo.Sedes s ON s.Id = e.SedeId
WHERE NOT EXISTS
(
    SELECT 1
    FROM dbo.NegocioClientes nc
    WHERE nc.NegocioId = s.NegocioId
      AND nc.ClienteId = r.ClienteId
);
GO

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
                      WHEN r.RolNegocio = 2 AND m.Codigo IN (N'DASHBOARD', N'RESERVAS', N'PAGOS', N'CLIENTES') THEN 1
                      WHEN r.RolNegocio = 3 AND m.Codigo IN (N'DASHBOARD', N'RESERVAS', N'CLIENTES') THEN 1
                      WHEN r.RolNegocio = 4 AND m.Codigo IN (N'DASHBOARD', N'PAGOS', N'COMPROBANTES') THEN 1
                      WHEN r.RolNegocio = 5 AND m.Codigo IN (N'DASHBOARD', N'RESERVAS', N'ESPACIOS') THEN 1
                      ELSE 0 END AS BIT),
            CAST(CASE WHEN r.RolNegocio = 1 THEN 1
                      WHEN r.RolNegocio = 2 AND m.Codigo IN (N'RESERVAS', N'PAGOS', N'CLIENTES') THEN 1
                      WHEN r.RolNegocio = 3 AND m.Codigo IN (N'RESERVAS', N'CLIENTES') THEN 1
                      WHEN r.RolNegocio = 4 AND m.Codigo IN (N'PAGOS', N'COMPROBANTES') THEN 1
                      ELSE 0 END AS BIT),
            CAST(CASE WHEN r.RolNegocio = 1 THEN 1
                      WHEN r.RolNegocio = 2 AND m.Codigo IN (N'RESERVAS', N'PAGOS', N'CLIENTES') THEN 1
                      WHEN r.RolNegocio = 3 AND m.Codigo IN (N'RESERVAS', N'CLIENTES') THEN 1
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

CREATE OR ALTER PROCEDURE dbo.Sp_Combos_Clientes
    @NegocioId INT
AS
BEGIN
    SET NOCOUNT ON;
    BEGIN TRY
        SELECT c.Id, CONCAT(c.NombresORazonSocial, N' (', c.NumeroDocumento, N')')
        FROM dbo.Clientes c
        INNER JOIN dbo.NegocioClientes nc ON nc.ClienteId = c.Id
        WHERE nc.NegocioId = @NegocioId
          AND nc.Activo = 1
          AND c.Activo = 1
        ORDER BY c.NombresORazonSocial;
    END TRY
    BEGIN CATCH
        DECLARE @ErrorMessage NVARCHAR(4000), @ErrorSeverity INT, @ErrorState INT;
        SELECT @ErrorMessage = ERROR_MESSAGE(), @ErrorSeverity = ERROR_SEVERITY(), @ErrorState = ERROR_STATE();
        RAISERROR (@ErrorMessage, @ErrorSeverity, @ErrorState);
    END CATCH
END
GO

CREATE OR ALTER PROCEDURE dbo.Sp_Clientes_Listar
    @NegocioId INT
AS
BEGIN
    SET NOCOUNT ON;
    BEGIN TRY
        SELECT
            c.Id,
            c.NombresORazonSocial,
            c.TipoDocumento,
            c.NumeroDocumento,
            c.Telefono,
            c.Correo,
            c.Activo
        FROM dbo.Clientes c
        INNER JOIN dbo.NegocioClientes nc ON nc.ClienteId = c.Id
        WHERE nc.NegocioId = @NegocioId
          AND nc.Activo = 1
        ORDER BY c.NombresORazonSocial;
    END TRY
    BEGIN CATCH
        DECLARE @ErrorMessage NVARCHAR(4000), @ErrorSeverity INT, @ErrorState INT;
        SELECT @ErrorMessage = ERROR_MESSAGE(), @ErrorSeverity = ERROR_SEVERITY(), @ErrorState = ERROR_STATE();
        RAISERROR (@ErrorMessage, @ErrorSeverity, @ErrorState);
    END CATCH
END
GO

CREATE OR ALTER PROCEDURE dbo.Sp_Clientes_ObtenerPorId
    @NegocioId INT,
    @Id INT
AS
BEGIN
    SET NOCOUNT ON;
    BEGIN TRY
        SELECT
            c.Id,
            c.NombresORazonSocial,
            c.TipoDocumento,
            c.NumeroDocumento,
            c.Telefono,
            c.Correo,
            c.DireccionFiscal,
            c.Activo
        FROM dbo.Clientes c
        INNER JOIN dbo.NegocioClientes nc ON nc.ClienteId = c.Id
        WHERE nc.NegocioId = @NegocioId
          AND nc.Activo = 1
          AND c.Id = @Id;
    END TRY
    BEGIN CATCH
        DECLARE @ErrorMessage NVARCHAR(4000), @ErrorSeverity INT, @ErrorState INT;
        SELECT @ErrorMessage = ERROR_MESSAGE(), @ErrorSeverity = ERROR_SEVERITY(), @ErrorState = ERROR_STATE();
        RAISERROR (@ErrorMessage, @ErrorSeverity, @ErrorState);
    END CATCH
END
GO

CREATE OR ALTER PROCEDURE dbo.Sp_Clientes_Crear
    @NegocioId INT,
    @NombresORazonSocial NVARCHAR(200),
    @TipoDocumento NVARCHAR(20),
    @NumeroDocumento NVARCHAR(20),
    @Telefono NVARCHAR(20) = NULL,
    @Correo NVARCHAR(200) = NULL,
    @DireccionFiscal NVARCHAR(250) = NULL,
    @Activo BIT,
    @Usuario NVARCHAR(200)
AS
BEGIN
    SET NOCOUNT ON;
    BEGIN TRY
        DECLARE @NumeroDocumentoNormalizado NVARCHAR(20);
        SET @NumeroDocumentoNormalizado = NULLIF(LTRIM(RTRIM(@NumeroDocumento)), N'');
        SET @NumeroDocumento = COALESCE(@NumeroDocumentoNormalizado, N'');

        IF @NumeroDocumentoNormalizado IS NOT NULL
           AND EXISTS
           (
               SELECT 1
               FROM dbo.Clientes c
               INNER JOIN dbo.NegocioClientes nc ON nc.ClienteId = c.Id
               WHERE nc.NegocioId = @NegocioId
                 AND nc.Activo = 1
                 AND c.Activo = 1
                 AND LTRIM(RTRIM(c.NumeroDocumento)) = @NumeroDocumentoNormalizado
           )
            RAISERROR('Cliente ya se encuentra registrado.', 16, 1);

        BEGIN TRANSACTION;

        INSERT INTO dbo.Clientes
        (
            NombresORazonSocial, TipoDocumento, NumeroDocumento, Telefono,
            Correo, DireccionFiscal, Activo, FechaCreacion, UsuarioCreacion
        )
        VALUES
        (
            @NombresORazonSocial, @TipoDocumento, @NumeroDocumento, @Telefono,
            @Correo, @DireccionFiscal, @Activo, SYSUTCDATETIME(), @Usuario
        );

        DECLARE @Id INT;
        SET @Id = SCOPE_IDENTITY();

        INSERT INTO dbo.NegocioClientes (NegocioId, ClienteId, Activo, FechaRegistro, UsuarioCreacion)
        VALUES (@NegocioId, @Id, 1, SYSUTCDATETIME(), @Usuario);

        DECLARE @EntidadIdAudit NVARCHAR(80);
        SET @EntidadIdAudit = CONVERT(NVARCHAR(80), @Id);
        EXEC dbo.Sp_Auditoria_Registrar @NegocioId = @NegocioId, @Modulo = N'CLIENTES', @Accion = N'CREATE', @Entidad = N'Cliente', @EntidadId = @EntidadIdAudit, @Usuario = @Usuario, @DetalleJson = NULL;

        COMMIT TRANSACTION;

        SELECT @Id;
    END TRY
    BEGIN CATCH
        IF XACT_STATE() <> 0
            ROLLBACK TRANSACTION;

        DECLARE @ErrorMessage NVARCHAR(4000), @ErrorSeverity INT, @ErrorState INT;
        SELECT @ErrorMessage = ERROR_MESSAGE(), @ErrorSeverity = ERROR_SEVERITY(), @ErrorState = ERROR_STATE();
        RAISERROR (@ErrorMessage, @ErrorSeverity, @ErrorState);
    END CATCH
END
GO

CREATE OR ALTER PROCEDURE dbo.Sp_Clientes_Actualizar
    @Id INT,
    @NegocioId INT,
    @NombresORazonSocial NVARCHAR(200),
    @TipoDocumento NVARCHAR(20),
    @NumeroDocumento NVARCHAR(20),
    @Telefono NVARCHAR(20) = NULL,
    @Correo NVARCHAR(200) = NULL,
    @DireccionFiscal NVARCHAR(250) = NULL,
    @Activo BIT,
    @Usuario NVARCHAR(200)
AS
BEGIN
    SET NOCOUNT ON;
    BEGIN TRY
        DECLARE @NumeroDocumentoNormalizado NVARCHAR(20);
        SET @NumeroDocumentoNormalizado = NULLIF(LTRIM(RTRIM(@NumeroDocumento)), N'');
        SET @NumeroDocumento = COALESCE(@NumeroDocumentoNormalizado, N'');

        IF @NumeroDocumentoNormalizado IS NOT NULL
           AND EXISTS
           (
               SELECT 1
               FROM dbo.Clientes c
               INNER JOIN dbo.NegocioClientes nc ON nc.ClienteId = c.Id
               WHERE nc.NegocioId = @NegocioId
                 AND nc.Activo = 1
                 AND c.Activo = 1
                 AND c.Id <> @Id
                 AND LTRIM(RTRIM(c.NumeroDocumento)) = @NumeroDocumentoNormalizado
           )
            RAISERROR('Cliente ya se encuentra registrado.', 16, 1);

        UPDATE c
        SET
            c.NombresORazonSocial = @NombresORazonSocial,
            c.TipoDocumento = @TipoDocumento,
            c.NumeroDocumento = @NumeroDocumento,
            c.Telefono = @Telefono,
            c.Correo = @Correo,
            c.DireccionFiscal = @DireccionFiscal,
            c.Activo = @Activo,
            c.FechaActualizacion = SYSUTCDATETIME(),
            c.UsuarioActualizacion = @Usuario
        FROM dbo.Clientes c
        INNER JOIN dbo.NegocioClientes nc ON nc.ClienteId = c.Id
        WHERE c.Id = @Id
          AND nc.NegocioId = @NegocioId
          AND nc.Activo = 1;

        IF @@ROWCOUNT = 0
            RAISERROR('No se encontro el cliente para actualizar en el negocio.', 16, 1);

        DECLARE @EntidadIdAudit NVARCHAR(80);
        SET @EntidadIdAudit = CONVERT(NVARCHAR(80), @Id);
        EXEC dbo.Sp_Auditoria_Registrar @NegocioId = @NegocioId, @Modulo = N'CLIENTES', @Accion = N'EDIT', @Entidad = N'Cliente', @EntidadId = @EntidadIdAudit, @Usuario = @Usuario, @DetalleJson = NULL;
    END TRY
    BEGIN CATCH
        DECLARE @ErrorMessage NVARCHAR(4000), @ErrorSeverity INT, @ErrorState INT;
        SELECT @ErrorMessage = ERROR_MESSAGE(), @ErrorSeverity = ERROR_SEVERITY(), @ErrorState = ERROR_STATE();
        RAISERROR (@ErrorMessage, @ErrorSeverity, @ErrorState);
    END CATCH
END
GO

CREATE OR ALTER PROCEDURE dbo.Sp_Clientes_Eliminar
    @NegocioId INT,
    @Id INT,
    @Usuario NVARCHAR(200)
AS
BEGIN
    SET NOCOUNT ON;
    BEGIN TRY
        UPDATE nc
        SET nc.Activo = 0
        FROM dbo.NegocioClientes nc
        WHERE nc.NegocioId = @NegocioId
          AND nc.ClienteId = @Id
          AND nc.Activo = 1;

        IF @@ROWCOUNT = 0
            RAISERROR('No se encontro el cliente para eliminar en el negocio.', 16, 1);

        IF NOT EXISTS (SELECT 1 FROM dbo.NegocioClientes WHERE ClienteId = @Id AND Activo = 1)
        BEGIN
            UPDATE dbo.Clientes
            SET Activo = 0,
                FechaActualizacion = SYSUTCDATETIME(),
                UsuarioActualizacion = @Usuario
            WHERE Id = @Id;
        END;

        DECLARE @EntidadIdAudit NVARCHAR(80);
        SET @EntidadIdAudit = CONVERT(NVARCHAR(80), @Id);
        EXEC dbo.Sp_Auditoria_Registrar @NegocioId = @NegocioId, @Modulo = N'CLIENTES', @Accion = N'DELETE', @Entidad = N'Cliente', @EntidadId = @EntidadIdAudit, @Usuario = @Usuario, @DetalleJson = NULL;
    END TRY
    BEGIN CATCH
        DECLARE @ErrorMessage NVARCHAR(4000), @ErrorSeverity INT, @ErrorState INT;
        SELECT @ErrorMessage = ERROR_MESSAGE(), @ErrorSeverity = ERROR_SEVERITY(), @ErrorState = ERROR_STATE();
        RAISERROR (@ErrorMessage, @ErrorSeverity, @ErrorState);
    END CATCH
END
GO
