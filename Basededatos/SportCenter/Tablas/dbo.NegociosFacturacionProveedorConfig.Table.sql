USE [DbSportCenter]
GO
/****** Object:  Table [dbo].[NegociosFacturacionProveedorConfig]    Script Date: 05/05/2026 ******/
-- Firma: Codex - 05/05/2026 | Crea configuracion por negocio/proveedor/ambiente para emision electronica.
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE TABLE [dbo].[NegociosFacturacionProveedorConfig](
    [Id] [int] IDENTITY(1,1) NOT NULL,
    [NegocioId] [int] NOT NULL,
    [ProveedorId] [int] NOT NULL,
    [Ambiente] [nvarchar](15) NOT NULL,
    [BaseUrl] [nvarchar](500) NOT NULL,
    [ApiVersion] [nvarchar](20) NULL,
    [TimeoutSegundos] [int] NOT NULL,
    [EsDefault] [bit] NOT NULL,
    [Activo] [bit] NOT NULL,
    [FechaRegistro] [datetime2](7) NOT NULL,
    [FechaActualizacion] [datetime2](7) NULL,
    [UsuarioActualizacion] [nvarchar](200) NULL,
 CONSTRAINT [PK_NegociosFacturacionProveedorConfig] PRIMARY KEY CLUSTERED
(
    [Id] ASC
)WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON) ON [PRIMARY]
) ON [PRIMARY]
GO
ALTER TABLE [dbo].[NegociosFacturacionProveedorConfig] ADD CONSTRAINT [DF_NegociosFacturacionProveedorConfig_TimeoutSegundos] DEFAULT ((30)) FOR [TimeoutSegundos]
GO
ALTER TABLE [dbo].[NegociosFacturacionProveedorConfig] ADD CONSTRAINT [DF_NegociosFacturacionProveedorConfig_EsDefault] DEFAULT ((0)) FOR [EsDefault]
GO
ALTER TABLE [dbo].[NegociosFacturacionProveedorConfig] ADD CONSTRAINT [DF_NegociosFacturacionProveedorConfig_Activo] DEFAULT ((1)) FOR [Activo]
GO
ALTER TABLE [dbo].[NegociosFacturacionProveedorConfig] ADD CONSTRAINT [DF_NegociosFacturacionProveedorConfig_FechaRegistro] DEFAULT (SYSUTCDATETIME()) FOR [FechaRegistro]
GO
ALTER TABLE [dbo].[NegociosFacturacionProveedorConfig] ADD CONSTRAINT [CK_NegociosFacturacionProveedorConfig_Ambiente]
CHECK ([Ambiente] IN (N'BETA', N'PRODUCCION'))
GO
ALTER TABLE [dbo].[NegociosFacturacionProveedorConfig] ADD CONSTRAINT [CK_NegociosFacturacionProveedorConfig_TimeoutSegundos]
CHECK ([TimeoutSegundos] BETWEEN 5 AND 300)
GO
ALTER TABLE [dbo].[NegociosFacturacionProveedorConfig]  WITH CHECK ADD CONSTRAINT [FK_NegociosFacturacionProveedorConfig_Negocios_NegocioId]
FOREIGN KEY([NegocioId]) REFERENCES [dbo].[Negocios] ([Id])
GO
ALTER TABLE [dbo].[NegociosFacturacionProveedorConfig] CHECK CONSTRAINT [FK_NegociosFacturacionProveedorConfig_Negocios_NegocioId]
GO
ALTER TABLE [dbo].[NegociosFacturacionProveedorConfig]  WITH CHECK ADD CONSTRAINT [FK_NegociosFacturacionProveedorConfig_FacturacionProveedores_ProveedorId]
FOREIGN KEY([ProveedorId]) REFERENCES [dbo].[FacturacionProveedores] ([Id])
GO
ALTER TABLE [dbo].[NegociosFacturacionProveedorConfig] CHECK CONSTRAINT [FK_NegociosFacturacionProveedorConfig_FacturacionProveedores_ProveedorId]
GO
CREATE UNIQUE NONCLUSTERED INDEX [UX_NegociosFacturacionProveedorConfig_Negocio_Proveedor_Ambiente]
ON [dbo].[NegociosFacturacionProveedorConfig]([NegocioId] ASC, [ProveedorId] ASC, [Ambiente] ASC)
GO
CREATE NONCLUSTERED INDEX [IX_NegociosFacturacionProveedorConfig_Negocio_Activo_Default]
ON [dbo].[NegociosFacturacionProveedorConfig]([NegocioId] ASC, [Activo] ASC, [EsDefault] ASC)
GO

