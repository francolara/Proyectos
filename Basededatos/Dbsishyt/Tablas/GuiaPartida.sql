-- =============================================
-- Author:        FRANCO LARA
-- Create date:   03/11/2026
-- Description:   Define GuiaPartida para la base inicial de HYT.
-- =============================================
USE [SistemaWebHYT];
GO
SET NOCOUNT ON;
IF OBJECT_ID(N'dbo.GuiaPartida', N'U') IS NULL
BEGIN
    CREATE TABLE dbo.GuiaPartida
    (
        PartidaId BIGINT IDENTITY NOT NULL CONSTRAINT PK_GuiaPartida PRIMARY KEY,
        GuiaId BIGINT NOT NULL,
        NumeroLinea INT NOT NULL,
        Posicion NVARCHAR(60) NULL,
        CodigoMaterialSAP NVARCHAR(60) NULL,
        Descripcion NVARCHAR(1500) NOT NULL,
        Unidad NVARCHAR(30) NOT NULL,
        Cantidad DECIMAL(19,6) NOT NULL,
        CONSTRAINT UQ_GuiaPartida_Linea UNIQUE (GuiaId, NumeroLinea),
        CONSTRAINT FK_GuiaPartida_Guia FOREIGN KEY (GuiaId) REFERENCES dbo.GuiaConformidad(GuiaId),
        CONSTRAINT CK_GuiaPartida_Cantidad CHECK (Cantidad > 0)
    );
END;
GO
