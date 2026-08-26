-- =============================================
-- Author:        FRANCO LARA
-- Create date:   16/06/2026
-- Description:   Maestro de tipos de comprobante segun catalogo SUNAT para compras y ventas.
-- =============================================

-- =============================================
-- Author:        FRANCO LARA
-- Create date:   20/06/2026
-- Description:   Agrega cuentas contables base separadas para compras y ventas.
-- =============================================

-- Firma: FRANCO LARA - 25/08/2026 | Reemplaza los identificadores de cuentas del maestro por codigos contables portables entre empresas.

IF OBJECT_ID(N'dbo.ADM_TipoComprobante', N'U') IS NULL
BEGIN
    CREATE TABLE dbo.ADM_TipoComprobante
    (
        IdTipoComprobante INT IDENTITY(1,1) NOT NULL CONSTRAINT PK_ADM_TipoComprobante PRIMARY KEY,
        CodigoTipoComprobante VARCHAR(3) NOT NULL,
        Descripcion NVARCHAR(150) NOT NULL,
        UsoCompras BIT NOT NULL CONSTRAINT DF_ADM_TipoComprobante_UsoCompras DEFAULT (0),
        UsoVentas BIT NOT NULL CONSTRAINT DF_ADM_TipoComprobante_UsoVentas DEFAULT (0),
        CodigoCuentaVentaSoles VARCHAR(20) NULL,
        CodigoCuentaVentaDolares VARCHAR(20) NULL,
        CodigoCuentaCompraSoles VARCHAR(20) NULL,
        CodigoCuentaCompraDolares VARCHAR(20) NULL,
        Estado BIT NOT NULL CONSTRAINT DF_ADM_TipoComprobante_Estado DEFAULT (1),
        FechaRegistro DATETIME2(0) NOT NULL CONSTRAINT DF_ADM_TipoComprobante_FechaRegistro DEFAULT (SYSDATETIME()),
        UsuarioRegistro NVARCHAR(450) NULL
    );

    ALTER TABLE dbo.ADM_TipoComprobante
        ADD CONSTRAINT UQ_ADM_TipoComprobante_Codigo UNIQUE (CodigoTipoComprobante);
END;
