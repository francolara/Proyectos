
GO

-- =============================================
-- Author:        FRANCO LARA
-- Create date:   10/06/2026
-- Firma Codex:   Base del modulo de boletines deportivos, agrega zona en ubigeo distritos y crea tabla principal de flyers/eventos.
-- =============================================

IF COL_LENGTH('dbo.UbigeoDistritos', 'Zona') IS NULL
BEGIN
    ALTER TABLE dbo.UbigeoDistritos ADD Zona NVARCHAR(20) NULL;
END;
GO

IF NOT EXISTS (SELECT 1 FROM sys.indexes WHERE name = N'IX_UbigeoDistritos_Zona' AND object_id = OBJECT_ID(N'dbo.UbigeoDistritos'))
BEGIN
    CREATE NONCLUSTERED INDEX IX_UbigeoDistritos_Zona
    ON dbo.UbigeoDistritos (Zona)
    WHERE Zona IS NOT NULL;
END;
GO

IF OBJECT_ID(N'dbo.BoletinesDeportivos', N'U') IS NULL
BEGIN
    CREATE TABLE dbo.BoletinesDeportivos
    (
        IdBoletin INT IDENTITY(1,1) NOT NULL,
        UsuarioId NVARCHAR(450) NOT NULL,
        PerfilPublicoId INT NULL,
        Titulo NVARCHAR(160) NULL,
        Descripcion NVARCHAR(500) NULL,
        ImagenUrl NVARCHAR(500) NOT NULL,
        FechaEvento DATE NOT NULL,
        CodigoUbigeo CHAR(6) NOT NULL,
        TipoRegistro CHAR(1) NOT NULL CONSTRAINT DF_BoletinesDeportivos_TipoRegistro DEFAULT ('U'),
        Activo BIT NOT NULL CONSTRAINT DF_BoletinesDeportivos_Activo DEFAULT (1),
        FechaCreacion DATETIME2(7) NOT NULL CONSTRAINT DF_BoletinesDeportivos_FechaCreacion DEFAULT (SYSDATETIME()),
        UsuarioCreacion NVARCHAR(120) NOT NULL,
        FechaActualizacion DATETIME2(7) NULL,
        UsuarioActualizacion NVARCHAR(120) NULL,
        CONSTRAINT PK_BoletinesDeportivos PRIMARY KEY CLUSTERED (IdBoletin),
        CONSTRAINT CK_BoletinesDeportivos_TipoRegistro CHECK (TipoRegistro IN ('U', 'A'))
    );
END;
GO

IF NOT EXISTS (SELECT 1 FROM sys.foreign_keys WHERE name = N'FK_BoletinesDeportivos_AspNetUsers_UsuarioId')
BEGIN
    ALTER TABLE dbo.BoletinesDeportivos
    WITH CHECK ADD CONSTRAINT FK_BoletinesDeportivos_AspNetUsers_UsuarioId
    FOREIGN KEY (UsuarioId)
    REFERENCES dbo.AspNetUsers (Id);
END;
GO

IF NOT EXISTS (SELECT 1 FROM sys.foreign_keys WHERE name = N'FK_BoletinesDeportivos_UsuariosPublicosPerfil_PerfilPublicoId')
BEGIN
    ALTER TABLE dbo.BoletinesDeportivos
    WITH CHECK ADD CONSTRAINT FK_BoletinesDeportivos_UsuariosPublicosPerfil_PerfilPublicoId
    FOREIGN KEY (PerfilPublicoId)
    REFERENCES dbo.UsuariosPublicosPerfil (Id);
END;
GO

IF NOT EXISTS (SELECT 1 FROM sys.foreign_keys WHERE name = N'FK_BoletinesDeportivos_UbigeoDistritos_CodigoUbigeo')
BEGIN
    ALTER TABLE dbo.BoletinesDeportivos
    WITH CHECK ADD CONSTRAINT FK_BoletinesDeportivos_UbigeoDistritos_CodigoUbigeo
    FOREIGN KEY (CodigoUbigeo)
    REFERENCES dbo.UbigeoDistritos (CodigoUbigeo);
END;
GO

IF NOT EXISTS (SELECT 1 FROM sys.indexes WHERE name = N'IX_BoletinesDeportivos_Activo_FechaEvento' AND object_id = OBJECT_ID(N'dbo.BoletinesDeportivos'))
BEGIN
    CREATE NONCLUSTERED INDEX IX_BoletinesDeportivos_Activo_FechaEvento
    ON dbo.BoletinesDeportivos (Activo, FechaEvento DESC, FechaCreacion DESC);
END;
GO

IF NOT EXISTS (SELECT 1 FROM sys.indexes WHERE name = N'IX_BoletinesDeportivos_UsuarioId_FechaCreacion' AND object_id = OBJECT_ID(N'dbo.BoletinesDeportivos'))
BEGIN
    CREATE NONCLUSTERED INDEX IX_BoletinesDeportivos_UsuarioId_FechaCreacion
    ON dbo.BoletinesDeportivos (UsuarioId, FechaCreacion DESC);
END;
GO

IF NOT EXISTS (SELECT 1 FROM sys.indexes WHERE name = N'IX_BoletinesDeportivos_CodigoUbigeo_FechaEvento' AND object_id = OBJECT_ID(N'dbo.BoletinesDeportivos'))
BEGIN
    CREATE NONCLUSTERED INDEX IX_BoletinesDeportivos_CodigoUbigeo_FechaEvento
    ON dbo.BoletinesDeportivos (CodigoUbigeo, FechaEvento DESC);
END;
GO
