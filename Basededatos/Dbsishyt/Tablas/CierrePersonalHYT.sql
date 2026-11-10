-- =============================================
-- Author:        FRANCO LARA
-- Create date:   09/11/2026
-- Description:   Define dbo.CierrePersonalHYT para las etapas 3 y 4 de Planillas.
-- =============================================
USE [SistemaWebHYT];
GO
IF OBJECT_ID(N'dbo.CierrePersonalHYT', N'U') IS NULL
BEGIN
    CREATE TABLE dbo.CierrePersonalHYT
    (
        Mes DATE NOT NULL CONSTRAINT PK_CierrePersonalHYT PRIMARY KEY,
        FechaCierre DATETIME2(3) NOT NULL CONSTRAINT DF_CierrePersonalHYT_Fecha DEFAULT SYSUTCDATETIME(),
        UsuarioId NVARCHAR(128) NOT NULL CONSTRAINT FK_CierrePersonalHYT_Usuario REFERENCES dbo.AspNetUsers(Id),
        CONSTRAINT CK_CierrePersonalHYT_Mes CHECK (DAY(Mes) = 1)
    );
END;
GO
