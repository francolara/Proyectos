-- =============================================
-- Author:        FRANCO LARA
-- Create date:   06/04/2026
-- Description:   Crea supermaestro TiposSueloSuperMaestro y relaciona TiposSuelo por TipoSueloSuperId.
-- =============================================

IF OBJECT_ID(N'dbo.TiposSueloSuperMaestro', N'U') IS NULL
BEGIN
    CREATE TABLE dbo.TiposSueloSuperMaestro
    (
        Id INT IDENTITY(1,1) NOT NULL CONSTRAINT PK_TiposSueloSuperMaestro PRIMARY KEY,
        Codigo NVARCHAR(20) NOT NULL,
        Nombre NVARCHAR(120) NOT NULL,
        Activo BIT NOT NULL CONSTRAINT DF_TiposSueloSuperMaestro_Activo DEFAULT (1),
        FechaCreacion DATETIME2(7) NOT NULL CONSTRAINT DF_TiposSueloSuperMaestro_FechaCreacion DEFAULT (SYSUTCDATETIME()),
        UsuarioCreacion NVARCHAR(200) NULL,
        FechaActualizacion DATETIME2(7) NULL,
        UsuarioActualizacion NVARCHAR(200) NULL
    );
END;

IF NOT EXISTS (
    SELECT 1
    FROM sys.indexes
    WHERE object_id = OBJECT_ID(N'dbo.TiposSueloSuperMaestro')
      AND name = N'UQ_TiposSueloSuperMaestro_Codigo'
)
BEGIN
    CREATE UNIQUE NONCLUSTERED INDEX UQ_TiposSueloSuperMaestro_Codigo
        ON dbo.TiposSueloSuperMaestro(Codigo);
END;

IF NOT EXISTS (
    SELECT 1
    FROM sys.indexes
    WHERE object_id = OBJECT_ID(N'dbo.TiposSueloSuperMaestro')
      AND name = N'UQ_TiposSueloSuperMaestro_Nombre'
)
BEGIN
    CREATE UNIQUE NONCLUSTERED INDEX UQ_TiposSueloSuperMaestro_Nombre
        ON dbo.TiposSueloSuperMaestro(Nombre);
END;

MERGE dbo.TiposSueloSuperMaestro AS destino
USING (
    VALUES
        (N'GRASS_SINT', N'Grass Sintetico'),
        (N'GRASS_NAT', N'Grass Natural'),
        (N'CEMENTO', N'Cemento'),
        (N'LOZA', N'Loza'),
        (N'PARQUET', N'Parquet'),
        (N'ARENA', N'Arena'),
        (N'TIERRA', N'Tierra')
) AS fuente (Codigo, Nombre)
ON destino.Codigo = fuente.Codigo
WHEN NOT MATCHED BY TARGET THEN
    INSERT (Codigo, Nombre, Activo, FechaCreacion, UsuarioCreacion)
    VALUES (fuente.Codigo, fuente.Nombre, 1, SYSUTCDATETIME(), N'sistema')
WHEN MATCHED THEN
    UPDATE SET
        destino.Nombre = fuente.Nombre,
        destino.Activo = 1,
        destino.FechaActualizacion = SYSUTCDATETIME(),
        destino.UsuarioActualizacion = N'sistema';

IF COL_LENGTH(N'dbo.TiposSuelo', N'TipoSueloSuperId') IS NULL
BEGIN
    ALTER TABLE dbo.TiposSuelo
        ADD TipoSueloSuperId INT NULL;
END;

UPDATE ts
SET ts.TipoSueloSuperId = tsm.Id,
    ts.Nombre = tsm.Nombre
FROM dbo.TiposSuelo ts
INNER JOIN dbo.TiposSueloSuperMaestro tsm
    ON UPPER(LTRIM(RTRIM(ts.Nombre))) = UPPER(LTRIM(RTRIM(tsm.Nombre)))
WHERE ts.TipoSueloSuperId IS NULL;

IF NOT EXISTS (
    SELECT 1
    FROM sys.foreign_keys
    WHERE name = N'FK_TiposSuelo_TiposSueloSuperMaestro_TipoSueloSuperId'
      AND parent_object_id = OBJECT_ID(N'dbo.TiposSuelo')
)
BEGIN
    ALTER TABLE dbo.TiposSuelo WITH CHECK
        ADD CONSTRAINT FK_TiposSuelo_TiposSueloSuperMaestro_TipoSueloSuperId
        FOREIGN KEY (TipoSueloSuperId)
        REFERENCES dbo.TiposSueloSuperMaestro(Id);
END;

IF NOT EXISTS (
    SELECT 1
    FROM sys.indexes
    WHERE name = N'UX_TiposSuelo_Negocio_TipoSueloSuperId'
      AND object_id = OBJECT_ID(N'dbo.TiposSuelo')
)
BEGIN
    CREATE UNIQUE NONCLUSTERED INDEX UX_TiposSuelo_Negocio_TipoSueloSuperId
        ON dbo.TiposSuelo(NegocioId, TipoSueloSuperId)
        WHERE NegocioId IS NOT NULL AND TipoSueloSuperId IS NOT NULL;
END;