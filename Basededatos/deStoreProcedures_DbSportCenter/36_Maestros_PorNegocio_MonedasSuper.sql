-- =============================================
-- Author:        FRANCO LARA
-- Create date:   02/04/2026
-- Description:   Maestros por negocio + supermaestro de monedas LATAM + ajustes de SP.
-- Firma:         Codex - 02/04/2026 | Agrega MonedasSuperMaestro, auditoria en catalogos, alcance por negocio y contratos SP para Maestros/Combos/Pagos/Espacios. Corrige auditoria MAESTROS (sin CONVERT en EXEC) y agrega columnas de auditoria faltantes en Negocios.
-- =============================================

IF OBJECT_ID(N'dbo.MonedasSuperMaestro', N'U') IS NULL
BEGIN
    CREATE TABLE dbo.MonedasSuperMaestro
    (
        Id INT IDENTITY(1,1) NOT NULL PRIMARY KEY,
        Codigo NVARCHAR(10) NOT NULL,
        Nombre NVARCHAR(120) NOT NULL,
        Simbolo NVARCHAR(10) NOT NULL,
        Activo BIT NOT NULL CONSTRAINT DF_MonedasSuperMaestro_Activo DEFAULT (1),
        FechaCreacion DATETIME2 NOT NULL CONSTRAINT DF_MonedasSuperMaestro_FechaCreacion DEFAULT (SYSUTCDATETIME()),
        UsuarioCreacion NVARCHAR(200) NULL,
        FechaActualizacion DATETIME2 NULL,
        UsuarioActualizacion NVARCHAR(200) NULL
    );
END;
GO

IF NOT EXISTS (SELECT 1 FROM sys.indexes WHERE object_id = OBJECT_ID(N'dbo.MonedasSuperMaestro') AND name = N'UX_MonedasSuperMaestro_Codigo')
    CREATE UNIQUE INDEX UX_MonedasSuperMaestro_Codigo ON dbo.MonedasSuperMaestro(Codigo);
GO

;WITH SeedMonedas AS
(
    SELECT N'ARS' Codigo, N'Peso argentino' Nombre, N'$' Simbolo UNION ALL
    SELECT N'BOB', N'Boliviano', N'Bs' UNION ALL
    SELECT N'BRL', N'Real brasileño', N'R$' UNION ALL
    SELECT N'CLP', N'Peso chileno', N'$' UNION ALL
    SELECT N'COP', N'Peso colombiano', N'$' UNION ALL
    SELECT N'CRC', N'Colon costarricense', N'CRC' UNION ALL
    SELECT N'DOP', N'Peso dominicano', N'RD$' UNION ALL
    SELECT N'GTQ', N'Quetzal', N'Q' UNION ALL
    SELECT N'HNL', N'Lempira', N'L' UNION ALL
    SELECT N'MXN', N'Peso mexicano', N'$' UNION ALL
    SELECT N'NIO', N'Cordoba', N'C$' UNION ALL
    SELECT N'PAB', N'Balboa', N'B/.' UNION ALL
    SELECT N'PEN', N'Sol peruano', N'S/' UNION ALL
    SELECT N'PYG', N'Guarani', N'Gs' UNION ALL
    SELECT N'UYU', N'Peso uruguayo', N'$U' UNION ALL
    SELECT N'VES', N'Bolivar', N'Bs.' UNION ALL
    SELECT N'USD', N'Dolar estadounidense', N'$'
)
MERGE dbo.MonedasSuperMaestro AS tgt
USING SeedMonedas AS src
ON tgt.Codigo = src.Codigo
WHEN MATCHED THEN UPDATE SET Nombre = src.Nombre, Simbolo = src.Simbolo, Activo = 1, FechaActualizacion = SYSUTCDATETIME(), UsuarioActualizacion = N'seed'
WHEN NOT MATCHED THEN INSERT (Codigo, Nombre, Simbolo, Activo, UsuarioCreacion) VALUES (src.Codigo, src.Nombre, src.Simbolo, 1, N'seed');
GO

IF COL_LENGTH('dbo.Monedas', 'NegocioId') IS NULL ALTER TABLE dbo.Monedas ADD NegocioId INT NULL;
IF COL_LENGTH('dbo.Monedas', 'MonedaSuperId') IS NULL ALTER TABLE dbo.Monedas ADD MonedaSuperId INT NULL;
IF COL_LENGTH('dbo.Monedas', 'FechaCreacion') IS NULL ALTER TABLE dbo.Monedas ADD FechaCreacion DATETIME2 NULL;
IF COL_LENGTH('dbo.Monedas', 'UsuarioCreacion') IS NULL ALTER TABLE dbo.Monedas ADD UsuarioCreacion NVARCHAR(200) NULL;
IF COL_LENGTH('dbo.Monedas', 'FechaActualizacion') IS NULL ALTER TABLE dbo.Monedas ADD FechaActualizacion DATETIME2 NULL;
IF COL_LENGTH('dbo.Monedas', 'UsuarioActualizacion') IS NULL ALTER TABLE dbo.Monedas ADD UsuarioActualizacion NVARCHAR(200) NULL;

IF COL_LENGTH('dbo.TiposSuelo', 'NegocioId') IS NULL ALTER TABLE dbo.TiposSuelo ADD NegocioId INT NULL;
IF COL_LENGTH('dbo.TiposSuelo', 'FechaCreacion') IS NULL ALTER TABLE dbo.TiposSuelo ADD FechaCreacion DATETIME2 NULL;
IF COL_LENGTH('dbo.TiposSuelo', 'UsuarioCreacion') IS NULL ALTER TABLE dbo.TiposSuelo ADD UsuarioCreacion NVARCHAR(200) NULL;
IF COL_LENGTH('dbo.TiposSuelo', 'FechaActualizacion') IS NULL ALTER TABLE dbo.TiposSuelo ADD FechaActualizacion DATETIME2 NULL;
IF COL_LENGTH('dbo.TiposSuelo', 'UsuarioActualizacion') IS NULL ALTER TABLE dbo.TiposSuelo ADD UsuarioActualizacion NVARCHAR(200) NULL;

IF COL_LENGTH('dbo.TiposDeporte', 'NegocioId') IS NULL ALTER TABLE dbo.TiposDeporte ADD NegocioId INT NULL;
IF COL_LENGTH('dbo.TiposDeporte', 'FechaCreacion') IS NULL ALTER TABLE dbo.TiposDeporte ADD FechaCreacion DATETIME2 NULL;
IF COL_LENGTH('dbo.TiposDeporte', 'UsuarioCreacion') IS NULL ALTER TABLE dbo.TiposDeporte ADD UsuarioCreacion NVARCHAR(200) NULL;
IF COL_LENGTH('dbo.TiposDeporte', 'FechaActualizacion') IS NULL ALTER TABLE dbo.TiposDeporte ADD FechaActualizacion DATETIME2 NULL;
IF COL_LENGTH('dbo.TiposDeporte', 'UsuarioActualizacion') IS NULL ALTER TABLE dbo.TiposDeporte ADD UsuarioActualizacion NVARCHAR(200) NULL;



IF COL_LENGTH('dbo.Negocios', 'FechaActualizacion') IS NULL ALTER TABLE dbo.Negocios ADD FechaActualizacion DATETIME2 NULL;
IF COL_LENGTH('dbo.Negocios', 'UsuarioActualizacion') IS NULL ALTER TABLE dbo.Negocios ADD UsuarioActualizacion NVARCHAR(200) NULL;
GO

IF EXISTS (SELECT 1 FROM sys.key_constraints WHERE name = N'UQ_FormasPago_Nombre')
    ALTER TABLE dbo.FormasPago DROP CONSTRAINT UQ_FormasPago_Nombre;
GO

-- Migracion minima a alcance por negocio (sin perder referencias actuales).
UPDATE td
SET td.NegocioId = s.NegocioId,
    td.FechaCreacion = COALESCE(td.FechaCreacion, SYSUTCDATETIME()),
    td.UsuarioCreacion = COALESCE(td.UsuarioCreacion, N'migracion_36')
FROM dbo.TiposDeporte td
INNER JOIN dbo.EspaciosDeportivos e ON e.TipoDeporteId = td.Id
INNER JOIN dbo.Sedes s ON s.Id = e.SedeId
WHERE td.NegocioId IS NULL;

UPDATE ts
SET ts.NegocioId = s.NegocioId,
    ts.FechaCreacion = COALESCE(ts.FechaCreacion, SYSUTCDATETIME()),
    ts.UsuarioCreacion = COALESCE(ts.UsuarioCreacion, N'migracion_36')
FROM dbo.TiposSuelo ts
INNER JOIN dbo.EspaciosDeportivos e ON e.TipoSueloId = ts.Id
INNER JOIN dbo.Sedes s ON s.Id = e.SedeId
WHERE ts.NegocioId IS NULL;

UPDATE fp
SET fp.NegocioId = s.NegocioId,
    fp.FechaCreacion = COALESCE(fp.FechaCreacion, SYSUTCDATETIME()),
    fp.UsuarioCreacion = COALESCE(fp.UsuarioCreacion, N'migracion_36')
FROM dbo.FormasPago fp
INNER JOIN dbo.Pagos p ON p.FormaPago = fp.Id
INNER JOIN dbo.Reservas r ON r.Id = p.ReservaId
INNER JOIN dbo.EspaciosDeportivos e ON e.Id = r.EspacioDeportivoId
INNER JOIN dbo.Sedes s ON s.Id = e.SedeId
WHERE fp.NegocioId IS NULL;
GO

-- Completa catalogos por cada negocio existente.
;WITH Negs AS (SELECT Id FROM dbo.Negocios)
INSERT INTO dbo.TiposDeporte (NegocioId, Nombre, Activo, FechaCreacion, UsuarioCreacion)
SELECT n.Id, x.Nombre, 1, SYSUTCDATETIME(), N'seed_36'
FROM Negs n
CROSS JOIN (SELECT N'Futbol' Nombre UNION ALL SELECT N'Voley' UNION ALL SELECT N'Basket' UNION ALL SELECT N'Fronton' UNION ALL SELECT N'Padel') x
WHERE NOT EXISTS (SELECT 1 FROM dbo.TiposDeporte td WHERE td.NegocioId = n.Id AND UPPER(LTRIM(RTRIM(td.Nombre))) = UPPER(LTRIM(RTRIM(x.Nombre))));

;WITH Negs AS (SELECT Id FROM dbo.Negocios)
INSERT INTO dbo.TiposSuelo (NegocioId, Nombre, Activo, FechaCreacion, UsuarioCreacion)
SELECT n.Id, x.Nombre, 1, SYSUTCDATETIME(), N'seed_36'
FROM Negs n
CROSS JOIN (SELECT N'Grass Sintetico' Nombre UNION ALL SELECT N'Grass Natural' UNION ALL SELECT N'Losa') x
WHERE NOT EXISTS (SELECT 1 FROM dbo.TiposSuelo ts WHERE ts.NegocioId = n.Id AND UPPER(LTRIM(RTRIM(ts.Nombre))) = UPPER(LTRIM(RTRIM(x.Nombre))));

;WITH Negs AS (SELECT Id FROM dbo.Negocios)
INSERT INTO dbo.FormasPago (NegocioId, Nombre, Activo, FechaCreacion, UsuarioCreacion)
SELECT n.Id, x.Nombre, 1, SYSUTCDATETIME(), N'seed_36'
FROM Negs n
CROSS JOIN (SELECT N'Efectivo' Nombre UNION ALL SELECT N'Yape' UNION ALL SELECT N'Plin' UNION ALL SELECT N'Transferencia' UNION ALL SELECT N'Tarjeta') x
WHERE NOT EXISTS (SELECT 1 FROM dbo.FormasPago fp WHERE fp.NegocioId = n.Id AND UPPER(LTRIM(RTRIM(fp.Nombre))) = UPPER(LTRIM(RTRIM(x.Nombre))));
GO

-- Monedas por negocio desde supermaestro.
UPDATE m
SET m.MonedaSuperId = sm.Id,
    m.FechaCreacion = COALESCE(m.FechaCreacion, SYSUTCDATETIME()),
    m.UsuarioCreacion = COALESCE(m.UsuarioCreacion, N'migracion_36')
FROM dbo.Monedas m
LEFT JOIN dbo.MonedasSuperMaestro sm ON sm.Codigo = UPPER(LTRIM(RTRIM(m.Codigo)))
WHERE m.MonedaSuperId IS NULL;

INSERT INTO dbo.Monedas (NegocioId, MonedaSuperId, Codigo, Nombre, Simbolo, Activo, FechaCreacion, UsuarioCreacion)
SELECT n.Id, sm.Id, sm.Codigo, sm.Nombre, sm.Simbolo, 1, SYSUTCDATETIME(), N'migracion_36'
FROM dbo.Negocios n
INNER JOIN dbo.MonedasSuperMaestro sm ON sm.Codigo = N'PEN'
WHERE NOT EXISTS (SELECT 1 FROM dbo.Monedas m WHERE m.NegocioId = n.Id);

UPDATE n
SET n.MonedaId = m.Id
FROM dbo.Negocios n
INNER JOIN dbo.Monedas m ON m.NegocioId = n.Id
WHERE n.MonedaId IS NULL;

UPDATE m
SET m.NegocioId = n.Id
FROM dbo.Monedas m
INNER JOIN dbo.Negocios n ON n.MonedaId = m.Id
WHERE m.NegocioId IS NULL;

IF EXISTS (SELECT 1 FROM dbo.Monedas WHERE NegocioId IS NULL)
BEGIN
    RAISERROR('Existen monedas sin NegocioId luego de la migracion. Revisa data manualmente.', 16, 1);
    RETURN;
END;

ALTER TABLE dbo.Monedas ALTER COLUMN NegocioId INT NOT NULL;
ALTER TABLE dbo.TiposDeporte ALTER COLUMN NegocioId INT NOT NULL;
ALTER TABLE dbo.TiposSuelo ALTER COLUMN NegocioId INT NOT NULL;
ALTER TABLE dbo.FormasPago ALTER COLUMN NegocioId INT NOT NULL;
GO

IF NOT EXISTS (SELECT 1 FROM sys.foreign_keys WHERE name = N'FK_Monedas_Negocios_NegocioId')
    ALTER TABLE dbo.Monedas ADD CONSTRAINT FK_Monedas_Negocios_NegocioId FOREIGN KEY (NegocioId) REFERENCES dbo.Negocios(Id);
IF NOT EXISTS (SELECT 1 FROM sys.foreign_keys WHERE name = N'FK_Monedas_MonedasSuperMaestro_MonedaSuperId')
    ALTER TABLE dbo.Monedas ADD CONSTRAINT FK_Monedas_MonedasSuperMaestro_MonedaSuperId FOREIGN KEY (MonedaSuperId) REFERENCES dbo.MonedasSuperMaestro(Id);
IF NOT EXISTS (SELECT 1 FROM sys.foreign_keys WHERE name = N'FK_TiposDeporte_Negocios_NegocioId')
    ALTER TABLE dbo.TiposDeporte ADD CONSTRAINT FK_TiposDeporte_Negocios_NegocioId FOREIGN KEY (NegocioId) REFERENCES dbo.Negocios(Id);
IF NOT EXISTS (SELECT 1 FROM sys.foreign_keys WHERE name = N'FK_TiposSuelo_Negocios_NegocioId')
    ALTER TABLE dbo.TiposSuelo ADD CONSTRAINT FK_TiposSuelo_Negocios_NegocioId FOREIGN KEY (NegocioId) REFERENCES dbo.Negocios(Id);
IF NOT EXISTS (SELECT 1 FROM sys.foreign_keys WHERE name = N'FK_FormasPago_Negocios_NegocioId')
    ALTER TABLE dbo.FormasPago ADD CONSTRAINT FK_FormasPago_Negocios_NegocioId FOREIGN KEY (NegocioId) REFERENCES dbo.Negocios(Id);
GO

IF NOT EXISTS (SELECT 1 FROM sys.indexes WHERE object_id = OBJECT_ID(N'dbo.Monedas') AND name = N'UX_Monedas_Negocio_MonedaSuper')
    CREATE UNIQUE INDEX UX_Monedas_Negocio_MonedaSuper ON dbo.Monedas(NegocioId, MonedaSuperId);
IF NOT EXISTS (SELECT 1 FROM sys.indexes WHERE object_id = OBJECT_ID(N'dbo.TiposDeporte') AND name = N'UX_TiposDeporte_Negocio_Nombre')
    CREATE UNIQUE INDEX UX_TiposDeporte_Negocio_Nombre ON dbo.TiposDeporte(NegocioId, Nombre);
IF NOT EXISTS (SELECT 1 FROM sys.indexes WHERE object_id = OBJECT_ID(N'dbo.TiposSuelo') AND name = N'UX_TiposSuelo_Negocio_Nombre')
    CREATE UNIQUE INDEX UX_TiposSuelo_Negocio_Nombre ON dbo.TiposSuelo(NegocioId, Nombre);
IF NOT EXISTS (SELECT 1 FROM sys.indexes WHERE object_id = OBJECT_ID(N'dbo.FormasPago') AND name = N'UX_FormasPago_Negocio_Nombre')
    CREATE UNIQUE INDEX UX_FormasPago_Negocio_Nombre ON dbo.FormasPago(NegocioId, Nombre);
GO

CREATE OR ALTER PROCEDURE dbo.Sp_Maestros_MonedasSuper_Listar
AS
BEGIN
    SET NOCOUNT ON;
    BEGIN TRY
        SELECT m.Id, CONCAT(m.Codigo, N' - ', m.Nombre, N' (', m.Simbolo, N')') AS Nombre
        FROM dbo.MonedasSuperMaestro m
        WHERE m.Activo = 1
        ORDER BY m.Codigo;
    END TRY
    BEGIN CATCH
        DECLARE @ErrorMessage NVARCHAR(4000), @ErrorSeverity INT, @ErrorState INT;
        SELECT @ErrorMessage = ERROR_MESSAGE(), @ErrorSeverity = ERROR_SEVERITY(), @ErrorState = ERROR_STATE();
        RAISERROR (@ErrorMessage, @ErrorSeverity, @ErrorState);
    END CATCH
END
GO

CREATE OR ALTER PROCEDURE dbo.Sp_Combos_Monedas
    @NegocioId INT
AS
BEGIN
    SET NOCOUNT ON;
    BEGIN TRY
        SELECT m.Id, CONCAT(m.Codigo, N' - ', m.Nombre, N' (', m.Simbolo, N')') AS Nombre
        FROM dbo.Monedas m
        WHERE m.NegocioId = @NegocioId
          AND m.Activo = 1
        ORDER BY m.Codigo;
    END TRY
    BEGIN CATCH
        DECLARE @ErrorMessage NVARCHAR(4000), @ErrorSeverity INT, @ErrorState INT;
        SELECT @ErrorMessage = ERROR_MESSAGE(), @ErrorSeverity = ERROR_SEVERITY(), @ErrorState = ERROR_STATE();
        RAISERROR (@ErrorMessage, @ErrorSeverity, @ErrorState);
    END CATCH
END
GO

CREATE OR ALTER PROCEDURE dbo.Sp_Combos_TiposDeporte
    @NegocioId INT
AS
BEGIN
    SET NOCOUNT ON;
    BEGIN TRY
        SELECT td.Id, td.Nombre
        FROM dbo.TiposDeporte td
        WHERE td.NegocioId = @NegocioId
          AND td.Activo = 1
        ORDER BY td.Nombre;
    END TRY
    BEGIN CATCH
        DECLARE @ErrorMessage NVARCHAR(4000), @ErrorSeverity INT, @ErrorState INT;
        SELECT @ErrorMessage = ERROR_MESSAGE(), @ErrorSeverity = ERROR_SEVERITY(), @ErrorState = ERROR_STATE();
        RAISERROR (@ErrorMessage, @ErrorSeverity, @ErrorState);
    END CATCH
END
GO

CREATE OR ALTER PROCEDURE dbo.Sp_Combos_TiposSuelo
    @NegocioId INT
AS
BEGIN
    SET NOCOUNT ON;
    BEGIN TRY
        SELECT ts.Id, ts.Nombre
        FROM dbo.TiposSuelo ts
        WHERE ts.NegocioId = @NegocioId
          AND ts.Activo = 1
        ORDER BY ts.Nombre;
    END TRY
    BEGIN CATCH
        DECLARE @ErrorMessage NVARCHAR(4000), @ErrorSeverity INT, @ErrorState INT;
        SELECT @ErrorMessage = ERROR_MESSAGE(), @ErrorSeverity = ERROR_SEVERITY(), @ErrorState = ERROR_STATE();
        RAISERROR (@ErrorMessage, @ErrorSeverity, @ErrorState);
    END CATCH
END
GO

CREATE OR ALTER PROCEDURE dbo.Sp_Combos_FormasPago
    @NegocioId INT
AS
BEGIN
    SET NOCOUNT ON;
    BEGIN TRY
        SELECT fp.Id, fp.Nombre
        FROM dbo.FormasPago fp
        WHERE fp.NegocioId = @NegocioId
          AND fp.Activo = 1
        ORDER BY fp.Nombre;
    END TRY
    BEGIN CATCH
        DECLARE @ErrorMessage NVARCHAR(4000), @ErrorSeverity INT, @ErrorState INT;
        SELECT @ErrorMessage = ERROR_MESSAGE(), @ErrorSeverity = ERROR_SEVERITY(), @ErrorState = ERROR_STATE();
        RAISERROR (@ErrorMessage, @ErrorSeverity, @ErrorState);
    END CATCH
END
GO

CREATE OR ALTER PROCEDURE dbo.Sp_Maestros_Monedas_Listar
    @NegocioId INT
AS
BEGIN
    SET NOCOUNT ON;
    BEGIN TRY
        SELECT m.Id, m.MonedaSuperId, m.Codigo, m.Nombre, m.Simbolo, m.Activo
        FROM dbo.Monedas m
        WHERE m.NegocioId = @NegocioId
        ORDER BY m.Nombre;
    END TRY
    BEGIN CATCH
        DECLARE @ErrorMessage NVARCHAR(4000), @ErrorSeverity INT, @ErrorState INT;
        SELECT @ErrorMessage = ERROR_MESSAGE(), @ErrorSeverity = ERROR_SEVERITY(), @ErrorState = ERROR_STATE();
        RAISERROR (@ErrorMessage, @ErrorSeverity, @ErrorState);
    END CATCH
END
GO

CREATE OR ALTER PROCEDURE dbo.Sp_Maestros_Monedas_Crear
    @NegocioId INT,
    @MonedaSuperId INT,
    @Activo BIT,
    @Usuario NVARCHAR(200)
AS
BEGIN
    SET NOCOUNT ON;
    BEGIN TRY
        IF NOT EXISTS (SELECT 1 FROM dbo.MonedasSuperMaestro WHERE Id = @MonedaSuperId AND Activo = 1)
            RAISERROR('La moneda del supermaestro no es valida.', 16, 1);
        IF EXISTS (SELECT 1 FROM dbo.Monedas WHERE NegocioId = @NegocioId AND MonedaSuperId = @MonedaSuperId)
            RAISERROR('La moneda ya esta registrada para este negocio.', 16, 1);

        INSERT INTO dbo.Monedas (NegocioId, MonedaSuperId, Codigo, Nombre, Simbolo, Activo, FechaCreacion, UsuarioCreacion)
        SELECT @NegocioId, ms.Id, ms.Codigo, ms.Nombre, ms.Simbolo, @Activo, SYSUTCDATETIME(), @Usuario
        FROM dbo.MonedasSuperMaestro ms
        WHERE ms.Id = @MonedaSuperId;

        DECLARE @Id INT = SCOPE_IDENTITY();
        EXEC dbo.Sp_Auditoria_Registrar @NegocioId = @NegocioId, @Modulo = N'MAESTROS', @Accion = N'CREATE', @Entidad = N'Moneda', @EntidadId = @Id, @Usuario = @Usuario, @DetalleJson = NULL;
        SELECT @Id;
    END TRY
    BEGIN CATCH
        DECLARE @ErrorMessage NVARCHAR(4000), @ErrorSeverity INT, @ErrorState INT;
        SELECT @ErrorMessage = ERROR_MESSAGE(), @ErrorSeverity = ERROR_SEVERITY(), @ErrorState = ERROR_STATE();
        RAISERROR (@ErrorMessage, @ErrorSeverity, @ErrorState);
    END CATCH
END
GO

CREATE OR ALTER PROCEDURE dbo.Sp_Maestros_Monedas_Actualizar
    @NegocioId INT,
    @Id INT,
    @Activo BIT,
    @Usuario NVARCHAR(200)
AS
BEGIN
    SET NOCOUNT ON;
    BEGIN TRY
        IF @Activo = 0 AND EXISTS (SELECT 1 FROM dbo.Negocios WHERE Id = @NegocioId AND MonedaId = @Id AND Activo = 1)
            RAISERROR('No se puede inactivar la moneda configurada en el club.', 16, 1);

        UPDATE dbo.Monedas
        SET Activo = @Activo,
            FechaActualizacion = SYSUTCDATETIME(),
            UsuarioActualizacion = @Usuario
        WHERE Id = @Id
          AND NegocioId = @NegocioId;

        IF @@ROWCOUNT = 0
            RAISERROR('No se encontro la moneda para actualizar.', 16, 1);

        EXEC dbo.Sp_Auditoria_Registrar @NegocioId = @NegocioId, @Modulo = N'MAESTROS', @Accion = N'EDIT', @Entidad = N'Moneda', @EntidadId = @Id, @Usuario = @Usuario, @DetalleJson = NULL;
    END TRY
    BEGIN CATCH
        DECLARE @ErrorMessage NVARCHAR(4000), @ErrorSeverity INT, @ErrorState INT;
        SELECT @ErrorMessage = ERROR_MESSAGE(), @ErrorSeverity = ERROR_SEVERITY(), @ErrorState = ERROR_STATE();
        RAISERROR (@ErrorMessage, @ErrorSeverity, @ErrorState);
    END CATCH
END
GO

CREATE OR ALTER PROCEDURE dbo.Sp_Maestros_Monedas_Eliminar
    @NegocioId INT,
    @Id INT,
    @Usuario NVARCHAR(200)
AS
BEGIN
    SET NOCOUNT ON;
    BEGIN TRY
        EXEC dbo.Sp_Maestros_Monedas_Actualizar @NegocioId = @NegocioId, @Id = @Id, @Activo = 0, @Usuario = @Usuario;
    END TRY
    BEGIN CATCH
        DECLARE @ErrorMessage NVARCHAR(4000), @ErrorSeverity INT, @ErrorState INT;
        SELECT @ErrorMessage = ERROR_MESSAGE(), @ErrorSeverity = ERROR_SEVERITY(), @ErrorState = ERROR_STATE();
        RAISERROR (@ErrorMessage, @ErrorSeverity, @ErrorState);
    END CATCH
END
GO

CREATE OR ALTER PROCEDURE dbo.Sp_Maestros_TiposSuelo_Listar
    @NegocioId INT
AS
BEGIN
    SET NOCOUNT ON;
    BEGIN TRY
        SELECT Id, Nombre, Activo FROM dbo.TiposSuelo WHERE NegocioId = @NegocioId ORDER BY Nombre;
    END TRY
    BEGIN CATCH
        DECLARE @ErrorMessage NVARCHAR(4000), @ErrorSeverity INT, @ErrorState INT;
        SELECT @ErrorMessage = ERROR_MESSAGE(), @ErrorSeverity = ERROR_SEVERITY(), @ErrorState = ERROR_STATE();
        RAISERROR (@ErrorMessage, @ErrorSeverity, @ErrorState);
    END CATCH
END
GO

CREATE OR ALTER PROCEDURE dbo.Sp_Maestros_TiposSuelo_Crear
    @NegocioId INT, @Nombre NVARCHAR(80), @Activo BIT, @Usuario NVARCHAR(200)
AS
BEGIN
    SET NOCOUNT ON;
    BEGIN TRY
        SET @Nombre = LTRIM(RTRIM(@Nombre));
        IF @Nombre = N'' RAISERROR('El nombre es obligatorio.',16,1);
        IF EXISTS (SELECT 1 FROM dbo.TiposSuelo WHERE NegocioId = @NegocioId AND UPPER(LTRIM(RTRIM(Nombre))) = UPPER(@Nombre))
            RAISERROR('Ya existe un tipo de suelo con ese nombre para el negocio.',16,1);
        INSERT INTO dbo.TiposSuelo(NegocioId,Nombre,Activo,FechaCreacion,UsuarioCreacion) VALUES(@NegocioId,@Nombre,@Activo,SYSUTCDATETIME(),@Usuario);
        DECLARE @Id INT = SCOPE_IDENTITY();
        EXEC dbo.Sp_Auditoria_Registrar @NegocioId=@NegocioId,@Modulo=N'MAESTROS',@Accion=N'CREATE',@Entidad=N'TipoSuelo',@EntidadId=@Id,@Usuario=@Usuario,@DetalleJson=NULL;
        SELECT @Id;
    END TRY
    BEGIN CATCH
        DECLARE @ErrorMessage NVARCHAR(4000), @ErrorSeverity INT, @ErrorState INT;
        SELECT @ErrorMessage = ERROR_MESSAGE(), @ErrorSeverity = ERROR_SEVERITY(), @ErrorState = ERROR_STATE();
        RAISERROR (@ErrorMessage, @ErrorSeverity, @ErrorState);
    END CATCH
END
GO

CREATE OR ALTER PROCEDURE dbo.Sp_Maestros_TiposSuelo_Actualizar
    @NegocioId INT, @Id INT, @Nombre NVARCHAR(80), @Activo BIT, @Usuario NVARCHAR(200)
AS
BEGIN
    SET NOCOUNT ON;
    BEGIN TRY
        SET @Nombre = LTRIM(RTRIM(@Nombre));
        UPDATE dbo.TiposSuelo SET Nombre=@Nombre,Activo=@Activo,FechaActualizacion=SYSUTCDATETIME(),UsuarioActualizacion=@Usuario WHERE Id=@Id AND NegocioId=@NegocioId;
        IF @@ROWCOUNT=0 RAISERROR('No se encontro el tipo de suelo para actualizar.',16,1);
        EXEC dbo.Sp_Auditoria_Registrar @NegocioId=@NegocioId,@Modulo=N'MAESTROS',@Accion=N'EDIT',@Entidad=N'TipoSuelo',@EntidadId=@Id,@Usuario=@Usuario,@DetalleJson=NULL;
    END TRY
    BEGIN CATCH
        DECLARE @ErrorMessage NVARCHAR(4000), @ErrorSeverity INT, @ErrorState INT;
        SELECT @ErrorMessage = ERROR_MESSAGE(), @ErrorSeverity = ERROR_SEVERITY(), @ErrorState = ERROR_STATE();
        RAISERROR (@ErrorMessage, @ErrorSeverity, @ErrorState);
    END CATCH
END
GO

CREATE OR ALTER PROCEDURE dbo.Sp_Maestros_TiposSuelo_Eliminar
    @NegocioId INT, @Id INT, @Usuario NVARCHAR(200)
AS
BEGIN
    SET NOCOUNT ON;
    BEGIN TRY
        UPDATE dbo.TiposSuelo SET Activo=0,FechaActualizacion=SYSUTCDATETIME(),UsuarioActualizacion=@Usuario WHERE Id=@Id AND NegocioId=@NegocioId;
        IF @@ROWCOUNT=0 RAISERROR('No se encontro el tipo de suelo para inactivar.',16,1);
        EXEC dbo.Sp_Auditoria_Registrar @NegocioId=@NegocioId,@Modulo=N'MAESTROS',@Accion=N'DELETE',@Entidad=N'TipoSuelo',@EntidadId=@Id,@Usuario=@Usuario,@DetalleJson=NULL;
    END TRY
    BEGIN CATCH
        DECLARE @ErrorMessage NVARCHAR(4000), @ErrorSeverity INT, @ErrorState INT;
        SELECT @ErrorMessage = ERROR_MESSAGE(), @ErrorSeverity = ERROR_SEVERITY(), @ErrorState = ERROR_STATE();
        RAISERROR (@ErrorMessage, @ErrorSeverity, @ErrorState);
    END CATCH
END
GO

CREATE OR ALTER PROCEDURE dbo.Sp_Maestros_TiposDeporte_Listar
    @NegocioId INT
AS
BEGIN
    SET NOCOUNT ON;
    BEGIN TRY
        SELECT Id, Nombre, Activo FROM dbo.TiposDeporte WHERE NegocioId=@NegocioId ORDER BY Nombre;
    END TRY
    BEGIN CATCH
        DECLARE @ErrorMessage NVARCHAR(4000), @ErrorSeverity INT, @ErrorState INT;
        SELECT @ErrorMessage = ERROR_MESSAGE(), @ErrorSeverity = ERROR_SEVERITY(), @ErrorState = ERROR_STATE();
        RAISERROR (@ErrorMessage, @ErrorSeverity, @ErrorState);
    END CATCH
END
GO

CREATE OR ALTER PROCEDURE dbo.Sp_Maestros_TiposDeporte_Crear
    @NegocioId INT, @Nombre NVARCHAR(80), @Activo BIT, @Usuario NVARCHAR(200)
AS
BEGIN
    SET NOCOUNT ON;
    BEGIN TRY
        SET @Nombre = LTRIM(RTRIM(@Nombre));
        IF @Nombre = N'' RAISERROR('El nombre es obligatorio.',16,1);
        IF EXISTS (SELECT 1 FROM dbo.TiposDeporte WHERE NegocioId=@NegocioId AND UPPER(LTRIM(RTRIM(Nombre)))=UPPER(@Nombre)) RAISERROR('Ya existe un tipo de deporte con ese nombre para el negocio.',16,1);
        INSERT INTO dbo.TiposDeporte(NegocioId,Nombre,Activo,FechaCreacion,UsuarioCreacion) VALUES(@NegocioId,@Nombre,@Activo,SYSUTCDATETIME(),@Usuario);
        DECLARE @Id INT = SCOPE_IDENTITY();
        EXEC dbo.Sp_Auditoria_Registrar @NegocioId=@NegocioId,@Modulo=N'MAESTROS',@Accion=N'CREATE',@Entidad=N'TipoDeporte',@EntidadId=@Id,@Usuario=@Usuario,@DetalleJson=NULL;
        SELECT @Id;
    END TRY
    BEGIN CATCH
        DECLARE @ErrorMessage NVARCHAR(4000), @ErrorSeverity INT, @ErrorState INT;
        SELECT @ErrorMessage = ERROR_MESSAGE(), @ErrorSeverity = ERROR_SEVERITY(), @ErrorState = ERROR_STATE();
        RAISERROR (@ErrorMessage, @ErrorSeverity, @ErrorState);
    END CATCH
END
GO

CREATE OR ALTER PROCEDURE dbo.Sp_Maestros_TiposDeporte_Actualizar
    @NegocioId INT, @Id INT, @Nombre NVARCHAR(80), @Activo BIT, @Usuario NVARCHAR(200)
AS
BEGIN
    SET NOCOUNT ON;
    BEGIN TRY
        UPDATE dbo.TiposDeporte SET Nombre=LTRIM(RTRIM(@Nombre)),Activo=@Activo,FechaActualizacion=SYSUTCDATETIME(),UsuarioActualizacion=@Usuario WHERE Id=@Id AND NegocioId=@NegocioId;
        IF @@ROWCOUNT=0 RAISERROR('No se encontro el tipo de deporte para actualizar.',16,1);
        EXEC dbo.Sp_Auditoria_Registrar @NegocioId=@NegocioId,@Modulo=N'MAESTROS',@Accion=N'EDIT',@Entidad=N'TipoDeporte',@EntidadId=@Id,@Usuario=@Usuario,@DetalleJson=NULL;
    END TRY
    BEGIN CATCH
        DECLARE @ErrorMessage NVARCHAR(4000), @ErrorSeverity INT, @ErrorState INT;
        SELECT @ErrorMessage = ERROR_MESSAGE(), @ErrorSeverity = ERROR_SEVERITY(), @ErrorState = ERROR_STATE();
        RAISERROR (@ErrorMessage, @ErrorSeverity, @ErrorState);
    END CATCH
END
GO

CREATE OR ALTER PROCEDURE dbo.Sp_Maestros_TiposDeporte_Eliminar
    @NegocioId INT, @Id INT, @Usuario NVARCHAR(200)
AS
BEGIN
    SET NOCOUNT ON;
    BEGIN TRY
        UPDATE dbo.TiposDeporte SET Activo=0,FechaActualizacion=SYSUTCDATETIME(),UsuarioActualizacion=@Usuario WHERE Id=@Id AND NegocioId=@NegocioId;
        IF @@ROWCOUNT=0 RAISERROR('No se encontro el tipo de deporte para inactivar.',16,1);
        EXEC dbo.Sp_Auditoria_Registrar @NegocioId=@NegocioId,@Modulo=N'MAESTROS',@Accion=N'DELETE',@Entidad=N'TipoDeporte',@EntidadId=@Id,@Usuario=@Usuario,@DetalleJson=NULL;
    END TRY
    BEGIN CATCH
        DECLARE @ErrorMessage NVARCHAR(4000), @ErrorSeverity INT, @ErrorState INT;
        SELECT @ErrorMessage = ERROR_MESSAGE(), @ErrorSeverity = ERROR_SEVERITY(), @ErrorState = ERROR_STATE();
        RAISERROR (@ErrorMessage, @ErrorSeverity, @ErrorState);
    END CATCH
END
GO

CREATE OR ALTER PROCEDURE dbo.Sp_Maestros_FormasPago_Listar
    @NegocioId INT
AS
BEGIN
    SET NOCOUNT ON;
    BEGIN TRY
        SELECT Id, Nombre, Activo FROM dbo.FormasPago WHERE NegocioId=@NegocioId ORDER BY Nombre;
    END TRY
    BEGIN CATCH
        DECLARE @ErrorMessage NVARCHAR(4000), @ErrorSeverity INT, @ErrorState INT;
        SELECT @ErrorMessage = ERROR_MESSAGE(), @ErrorSeverity = ERROR_SEVERITY(), @ErrorState = ERROR_STATE();
        RAISERROR (@ErrorMessage, @ErrorSeverity, @ErrorState);
    END CATCH
END
GO

CREATE OR ALTER PROCEDURE dbo.Sp_Maestros_FormasPago_Crear
    @NegocioId INT, @Nombre NVARCHAR(80), @Activo BIT, @Usuario NVARCHAR(200)
AS
BEGIN
    SET NOCOUNT ON;
    BEGIN TRY
        INSERT INTO dbo.FormasPago(NegocioId,Nombre,Activo,FechaCreacion,UsuarioCreacion) VALUES(@NegocioId,LTRIM(RTRIM(@Nombre)),@Activo,SYSUTCDATETIME(),@Usuario);
        DECLARE @Id INT = SCOPE_IDENTITY();
        EXEC dbo.Sp_Auditoria_Registrar @NegocioId=@NegocioId,@Modulo=N'MAESTROS',@Accion=N'CREATE',@Entidad=N'FormaPago',@EntidadId=@Id,@Usuario=@Usuario,@DetalleJson=NULL;
        SELECT @Id;
    END TRY
    BEGIN CATCH
        DECLARE @ErrorMessage NVARCHAR(4000), @ErrorSeverity INT, @ErrorState INT;
        SELECT @ErrorMessage = ERROR_MESSAGE(), @ErrorSeverity = ERROR_SEVERITY(), @ErrorState = ERROR_STATE();
        RAISERROR (@ErrorMessage, @ErrorSeverity, @ErrorState);
    END CATCH
END
GO

CREATE OR ALTER PROCEDURE dbo.Sp_Maestros_FormasPago_Actualizar
    @NegocioId INT, @Id INT, @Nombre NVARCHAR(80), @Activo BIT, @Usuario NVARCHAR(200)
AS
BEGIN
    SET NOCOUNT ON;
    BEGIN TRY
        UPDATE dbo.FormasPago SET Nombre=LTRIM(RTRIM(@Nombre)),Activo=@Activo,FechaActualizacion=SYSUTCDATETIME(),UsuarioActualizacion=@Usuario WHERE Id=@Id AND NegocioId=@NegocioId;
        IF @@ROWCOUNT=0 RAISERROR('No se encontro la forma de pago para actualizar.',16,1);
        EXEC dbo.Sp_Auditoria_Registrar @NegocioId=@NegocioId,@Modulo=N'MAESTROS',@Accion=N'EDIT',@Entidad=N'FormaPago',@EntidadId=@Id,@Usuario=@Usuario,@DetalleJson=NULL;
    END TRY
    BEGIN CATCH
        DECLARE @ErrorMessage NVARCHAR(4000), @ErrorSeverity INT, @ErrorState INT;
        SELECT @ErrorMessage = ERROR_MESSAGE(), @ErrorSeverity = ERROR_SEVERITY(), @ErrorState = ERROR_STATE();
        RAISERROR (@ErrorMessage, @ErrorSeverity, @ErrorState);
    END CATCH
END
GO

CREATE OR ALTER PROCEDURE dbo.Sp_Maestros_FormasPago_Eliminar
    @NegocioId INT, @Id INT, @Usuario NVARCHAR(200)
AS
BEGIN
    SET NOCOUNT ON;
    BEGIN TRY
        UPDATE dbo.FormasPago SET Activo=0,FechaActualizacion=SYSUTCDATETIME(),UsuarioActualizacion=@Usuario WHERE Id=@Id AND NegocioId=@NegocioId;
        IF @@ROWCOUNT=0 RAISERROR('No se encontro la forma de pago para inactivar.',16,1);
        EXEC dbo.Sp_Auditoria_Registrar @NegocioId=@NegocioId,@Modulo=N'MAESTROS',@Accion=N'DELETE',@Entidad=N'FormaPago',@EntidadId=@Id,@Usuario=@Usuario,@DetalleJson=NULL;
    END TRY
    BEGIN CATCH
        DECLARE @ErrorMessage NVARCHAR(4000), @ErrorSeverity INT, @ErrorState INT;
        SELECT @ErrorMessage = ERROR_MESSAGE(), @ErrorSeverity = ERROR_SEVERITY(), @ErrorState = ERROR_STATE();
        RAISERROR (@ErrorMessage, @ErrorSeverity, @ErrorState);
    END CATCH
END
GO

CREATE OR ALTER PROCEDURE dbo.Sp_ConfiguracionClub_Actualizar
    @NegocioId INT,
    @NombreComercial NVARCHAR(200),
    @RazonSocial NVARCHAR(200) = NULL,
    @TipoDocumentoFiscal NVARCHAR(20) = NULL,
    @NumeroDocumentoFiscal NVARCHAR(20) = NULL,
    @DireccionFiscal NVARCHAR(250) = NULL,
    @MonedaId INT,
    @Usuario NVARCHAR(200)
AS
BEGIN
    SET NOCOUNT ON;
    BEGIN TRY
        IF NOT EXISTS (SELECT 1 FROM dbo.Monedas WHERE Id = @MonedaId AND NegocioId = @NegocioId AND Activo = 1)
            RAISERROR('La moneda seleccionada no pertenece al negocio.', 16, 1);

        UPDATE dbo.Negocios
        SET NombreComercial = @NombreComercial,
            RazonSocial = @RazonSocial,
            TipoDocumentoFiscal = @TipoDocumentoFiscal,
            NumeroDocumentoFiscal = @NumeroDocumentoFiscal,
            DireccionFiscal = @DireccionFiscal,
            MonedaId = @MonedaId
        WHERE Id = @NegocioId;

        IF @@ROWCOUNT = 0
            RAISERROR('No se encontro el negocio para actualizar.', 16, 1);
    END TRY
    BEGIN CATCH
        DECLARE @ErrorMessage NVARCHAR(4000), @ErrorSeverity INT, @ErrorState INT;
        SELECT @ErrorMessage = ERROR_MESSAGE(), @ErrorSeverity = ERROR_SEVERITY(), @ErrorState = ERROR_STATE();
        RAISERROR (@ErrorMessage, @ErrorSeverity, @ErrorState);
    END CATCH
END
GO


