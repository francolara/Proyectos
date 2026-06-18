-- =============================================
-- Author:        FRANCO LARA
-- Create date:   16/06/2026
-- Description:   Maestro de tipos de comprobante segun catalogo SUNAT para compras y ventas.
-- =============================================

IF OBJECT_ID(N'dbo.ADM_TipoComprobante', N'U') IS NULL
BEGIN
    CREATE TABLE dbo.ADM_TipoComprobante
    (
        IdTipoComprobante INT IDENTITY(1,1) NOT NULL CONSTRAINT PK_ADM_TipoComprobante PRIMARY KEY,
        CodigoTipoComprobante VARCHAR(3) NOT NULL,
        Descripcion NVARCHAR(150) NOT NULL,
        UsoCompras BIT NOT NULL CONSTRAINT DF_ADM_TipoComprobante_UsoCompras DEFAULT (0),
        UsoVentas BIT NOT NULL CONSTRAINT DF_ADM_TipoComprobante_UsoVentas DEFAULT (0),
        Estado BIT NOT NULL CONSTRAINT DF_ADM_TipoComprobante_Estado DEFAULT (1),
        FechaRegistro DATETIME2(0) NOT NULL CONSTRAINT DF_ADM_TipoComprobante_FechaRegistro DEFAULT (SYSDATETIME()),
        UsuarioRegistro NVARCHAR(450) NULL
    );

    ALTER TABLE dbo.ADM_TipoComprobante
        ADD CONSTRAINT UQ_ADM_TipoComprobante_Codigo UNIQUE (CodigoTipoComprobante);
END;
