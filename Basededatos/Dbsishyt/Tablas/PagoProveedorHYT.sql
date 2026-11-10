-- =============================================
-- Author:        FRANCO LARA
-- Create date:   09/11/2026
-- Description:   Define dbo.PagoProveedorHYT para el modulo Planillas.
-- =============================================
USE [SistemaWebHYT];
GO
SET NOCOUNT ON;
IF OBJECT_ID(N'dbo.PagoProveedorHYT', N'U') IS NULL
BEGIN
    CREATE TABLE dbo.PagoProveedorHYT
    (
        PagoProveedorId BIGINT IDENTITY NOT NULL CONSTRAINT PK_PagoProveedorHYT PRIMARY KEY,
        CostoProveedorId BIGINT NOT NULL CONSTRAINT FK_PagoProveedor_Costo REFERENCES dbo.CostoProveedorPresupuesto(CostoProveedorId),
        Importe DECIMAL(18,2) NOT NULL CONSTRAINT CK_PagoProveedor_Importe CHECK (Importe > 0),
        FechaPago DATE NOT NULL,
        Referencia NVARCHAR(100) NOT NULL,
        OperacionId UNIQUEIDENTIFIER NOT NULL CONSTRAINT UQ_PagoProveedor_Operacion UNIQUE,
        FechaRegistro DATETIME2(3) NOT NULL CONSTRAINT DF_PagoProveedor_Fecha DEFAULT SYSUTCDATETIME(),
        UsuarioRegistro NVARCHAR(128) NOT NULL CONSTRAINT FK_PagoProveedor_Usuario REFERENCES dbo.AspNetUsers(Id)
    );
END;
GO
