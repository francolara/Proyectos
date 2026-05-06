USE [DbSportCenter]
GO
/****** Object:  Table [dbo].[FacturacionProveedores]    Script Date: 05/05/2026 ******/
-- Firma: Codex - 05/05/2026 | Crea catalogo de proveedores de facturacion electronica para configuracion multi-proveedor.
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE TABLE [dbo].[FacturacionProveedores](
    [Id] [int] IDENTITY(1,1) NOT NULL,
    [Codigo] [nvarchar](30) NOT NULL,
    [Nombre] [nvarchar](120) NOT NULL,
    [TipoAutenticacion] [nvarchar](20) NOT NULL,
    [Activo] [bit] NOT NULL,
    [FechaRegistro] [datetime2](7) NOT NULL,
    [FechaActualizacion] [datetime2](7) NULL,
    [UsuarioActualizacion] [nvarchar](200) NULL,
 CONSTRAINT [PK_FacturacionProveedores] PRIMARY KEY CLUSTERED
(
    [Id] ASC
)WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON) ON [PRIMARY],
 CONSTRAINT [UQ_FacturacionProveedores_Codigo] UNIQUE NONCLUSTERED
(
    [Codigo] ASC
)WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON) ON [PRIMARY]
) ON [PRIMARY]
GO
ALTER TABLE [dbo].[FacturacionProveedores] ADD CONSTRAINT [DF_FacturacionProveedores_Activo] DEFAULT ((1)) FOR [Activo]
GO
ALTER TABLE [dbo].[FacturacionProveedores] ADD CONSTRAINT [DF_FacturacionProveedores_FechaRegistro] DEFAULT (SYSUTCDATETIME()) FOR [FechaRegistro]
GO
ALTER TABLE [dbo].[FacturacionProveedores] ADD CONSTRAINT [CK_FacturacionProveedores_TipoAutenticacion]
CHECK ([TipoAutenticacion] IN (N'API_KEY', N'OAUTH2_CLIENT', N'USER_PASS', N'TOKEN_FIJO'))
GO

