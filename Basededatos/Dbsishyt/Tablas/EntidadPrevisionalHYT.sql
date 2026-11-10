-- =============================================
-- Author:        FRANCO LARA
-- Create date:   09/11/2026
-- Description:   EntidadPrevisionalHYT
-- =============================================
USE [SistemaWebHYT];
GO
IF OBJECT_ID(N'dbo.EntidadPrevisionalHYT', N'U') IS NULL
BEGIN
    CREATE TABLE dbo.EntidadPrevisionalHYT
    (
        Nombre NVARCHAR(100) NOT NULL CONSTRAINT PK_EntidadPrevisionalHYT PRIMARY KEY,
        Regimen VARCHAR(3) NOT NULL CONSTRAINT CK_EntidadPrevisionalHYT CHECK (Regimen IN ('AFP', 'ONP'))
    );
END;
GO
