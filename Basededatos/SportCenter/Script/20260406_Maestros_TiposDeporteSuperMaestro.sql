-- =============================================
-- Author:        FRANCO LARA
-- Create date:   06/04/2026
-- Description:   Crea supermaestro TiposDeporteSuperMaestro y relaciona TiposDeporte por TipoDeporteSuperId.
-- =============================================

IF OBJECT_ID(N'dbo.TiposDeporteSuperMaestro', N'U') IS NULL
BEGIN
    CREATE TABLE dbo.TiposDeporteSuperMaestro
    (
        Id INT IDENTITY(1,1) NOT NULL CONSTRAINT PK_TiposDeporteSuperMaestro PRIMARY KEY,
        Codigo NVARCHAR(20) NOT NULL,
        Nombre NVARCHAR(120) NOT NULL,
        Activo BIT NOT NULL CONSTRAINT DF_TiposDeporteSuperMaestro_Activo DEFAULT (1),
        FechaCreacion DATETIME2(7) NOT NULL CONSTRAINT DF_TiposDeporteSuperMaestro_FechaCreacion DEFAULT (SYSUTCDATETIME()),
        UsuarioCreacion NVARCHAR(200) NULL,
        FechaActualizacion DATETIME2(7) NULL,
        UsuarioActualizacion NVARCHAR(200) NULL
    );
END;

IF NOT EXISTS (
    SELECT 1
    FROM sys.indexes
    WHERE object_id = OBJECT_ID(N'dbo.TiposDeporteSuperMaestro')
      AND name = N'UQ_TiposDeporteSuperMaestro_Codigo'
)
BEGIN
    CREATE UNIQUE NONCLUSTERED INDEX UQ_TiposDeporteSuperMaestro_Codigo
        ON dbo.TiposDeporteSuperMaestro(Codigo);
END;

IF NOT EXISTS (
    SELECT 1
    FROM sys.indexes
    WHERE object_id = OBJECT_ID(N'dbo.TiposDeporteSuperMaestro')
      AND name = N'UQ_TiposDeporteSuperMaestro_Nombre'
)
BEGIN
    CREATE UNIQUE NONCLUSTERED INDEX UQ_TiposDeporteSuperMaestro_Nombre
        ON dbo.TiposDeporteSuperMaestro(Nombre);
END;

MERGE dbo.TiposDeporteSuperMaestro AS destino
USING (
    VALUES
        (N'FUTBOL', N'Futbol'),
        (N'FULBITO', N'Fulbito'),
        (N'VOLEY', N'Voley'),
        (N'BASQUET', N'Basquet'),
        (N'TENIS', N'Tenis'),
        (N'PADEL', N'Padel')
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

IF COL_LENGTH(N'dbo.TiposDeporte', N'TipoDeporteSuperId') IS NULL
BEGIN
    ALTER TABLE dbo.TiposDeporte
        ADD TipoDeporteSuperId INT NULL;
END;

UPDATE td
SET td.TipoDeporteSuperId = tdm.Id,
    td.Nombre = tdm.Nombre
FROM dbo.TiposDeporte td
INNER JOIN dbo.TiposDeporteSuperMaestro tdm
    ON UPPER(LTRIM(RTRIM(td.Nombre))) = UPPER(LTRIM(RTRIM(tdm.Nombre)))
WHERE td.TipoDeporteSuperId IS NULL;

IF NOT EXISTS (
    SELECT 1
    FROM sys.foreign_keys
    WHERE name = N'FK_TiposDeporte_TiposDeporteSuperMaestro_TipoDeporteSuperId'
      AND parent_object_id = OBJECT_ID(N'dbo.TiposDeporte')
)
BEGIN
    ALTER TABLE dbo.TiposDeporte WITH CHECK
        ADD CONSTRAINT FK_TiposDeporte_TiposDeporteSuperMaestro_TipoDeporteSuperId
        FOREIGN KEY (TipoDeporteSuperId)
        REFERENCES dbo.TiposDeporteSuperMaestro(Id);
END;

IF NOT EXISTS (
    SELECT 1
    FROM sys.indexes
    WHERE name = N'UX_TiposDeporte_Negocio_TipoDeporteSuperId'
      AND object_id = OBJECT_ID(N'dbo.TiposDeporte')
)
BEGIN
    CREATE UNIQUE NONCLUSTERED INDEX UX_TiposDeporte_Negocio_TipoDeporteSuperId
        ON dbo.TiposDeporte(NegocioId, TipoDeporteSuperId)
        WHERE NegocioId IS NOT NULL AND TipoDeporteSuperId IS NOT NULL;
END;