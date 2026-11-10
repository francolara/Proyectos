-- =============================================
-- Author:        FRANCO LARA
-- Create date:   09/11/2026
-- Description:   Define dbo.CostoProveedorPresupuesto para el modulo Planillas.
-- =============================================
USE [SistemaWebHYT];
GO
SET NOCOUNT ON;
IF OBJECT_ID(N'dbo.CostoProveedorPresupuesto', N'U') IS NULL
BEGIN
    CREATE TABLE dbo.CostoProveedorPresupuesto
    (
        CostoProveedorId BIGINT IDENTITY NOT NULL CONSTRAINT PK_CostoProveedorPresupuesto PRIMARY KEY,
        PresupuestoId BIGINT NOT NULL CONSTRAINT FK_CostoProveedor_Presupuesto REFERENCES dbo.Presupuesto(PresupuestoId),
        ProveedorId BIGINT NOT NULL CONSTRAINT FK_CostoProveedor_Proveedor REFERENCES dbo.ProveedorHYT(ProveedorId),
        Descripcion NVARCHAR(300) NOT NULL,
        Importe DECIMAL(18,2) NOT NULL CONSTRAINT CK_CostoProveedor_Importe CHECK (Importe > 0),
        MonedaCodigo NVARCHAR(3) NOT NULL CONSTRAINT FK_CostoProveedor_Moneda REFERENCES dbo.Moneda(Codigo),
        DocumentoSustento NVARCHAR(100) NULL,
        OperacionId UNIQUEIDENTIFIER NOT NULL CONSTRAINT UQ_CostoProveedor_Operacion UNIQUE,
        FechaRegistro DATETIME2(3) NOT NULL CONSTRAINT DF_CostoProveedor_Fecha DEFAULT SYSUTCDATETIME(),
        UsuarioRegistro NVARCHAR(128) NOT NULL CONSTRAINT FK_CostoProveedor_Usuario REFERENCES dbo.AspNetUsers(Id)
    );
END;
GO
