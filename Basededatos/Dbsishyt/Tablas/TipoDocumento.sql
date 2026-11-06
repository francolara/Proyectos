-- =============================================
-- Author:        FRANCO LARA
-- Create date:   03/11/2026
-- Description:   Prepara el catalogo interno de tipos de documento
--                con codigo textual, sin identificador numerico.
-- =============================================
USE [SistemaWebHYT];
GO
SET NOCOUNT ON;

IF OBJECT_ID(N'dbo.TipoDocumento', N'U') IS NULL
BEGIN
    CREATE TABLE dbo.TipoDocumento
    (
        Codigo NVARCHAR(2) NOT NULL,
        Nombre NVARCHAR(40) NOT NULL,
        CONSTRAINT PK_TipoDocumento PRIMARY KEY (Codigo),
        CONSTRAINT CK_TipoDocumento_Codigo CHECK (Codigo IN (N'01', N'03'))
    );
END;
GO
