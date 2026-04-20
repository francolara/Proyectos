USE [DbSportCenter]
GO
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO

-- Firma: Codex - 19/04/2026 | ALTER TABLE incremental para agregar UsuariosPermitidos en Negocios con default 3.

IF COL_LENGTH('dbo.Negocios', 'UsuariosPermitidos') IS NULL
BEGIN
    ALTER TABLE dbo.Negocios
    ADD UsuariosPermitidos INT NOT NULL
        CONSTRAINT DF_Negocios_UsuariosPermitidos DEFAULT ((3));
END
GO
