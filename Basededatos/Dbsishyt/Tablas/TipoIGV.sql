-- =============================================
-- Author:        FRANCO LARA
-- Create date:   03/11/2026
-- Description:   Define tratamientos de precio CON IGV y SIN IGV
--                mediante codigo textual, sin identificador numerico.
-- =============================================
USE [SistemaWebHYT];
GO
SET NOCOUNT ON;

IF OBJECT_ID(N'dbo.TipoIGV', N'U') IS NULL
BEGIN
    CREATE TABLE dbo.TipoIGV
    (
        Codigo NVARCHAR(10) NOT NULL,
        Nombre NVARCHAR(20) NOT NULL,
        IncluyeIGV BIT NOT NULL,
        CONSTRAINT PK_TipoIGV PRIMARY KEY (Codigo),
        CONSTRAINT CK_TipoIGV_Tratamiento CHECK
        (
            (Codigo = N'CON_IGV' AND IncluyeIGV = 1)
            OR (Codigo = N'SIN_IGV' AND IncluyeIGV = 0)
        )
    );
END;
GO
