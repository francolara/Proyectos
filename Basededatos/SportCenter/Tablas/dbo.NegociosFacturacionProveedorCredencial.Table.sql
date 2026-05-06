USE [DbSportCenter]
GO
/****** Object:  Table [dbo].[NegociosFacturacionProveedorCredencial]    Script Date: 05/05/2026 ******/
-- Firma: Codex - 05/05/2026 | Crea almacenamiento cifrado de credenciales por negocio y proveedor.
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE TABLE [dbo].[NegociosFacturacionProveedorCredencial](
    [Id] [int] IDENTITY(1,1) NOT NULL,
    [NegocioProveedorConfigId] [int] NOT NULL,
    [TipoCredencial] [nvarchar](30) NOT NULL,
    [SecretoCifrado] [varbinary](max) NOT NULL,
    [KeyVersion] [nvarchar](20) NOT NULL,
    [ExpiraEn] [datetime2](7) NULL,
    [Scope] [nvarchar](200) NULL,
    [Activo] [bit] NOT NULL,
    [FechaRegistro] [datetime2](7) NOT NULL,
    [FechaActualizacion] [datetime2](7) NULL,
    [UsuarioActualizacion] [nvarchar](200) NULL,
 CONSTRAINT [PK_NegociosFacturacionProveedorCredencial] PRIMARY KEY CLUSTERED
(
    [Id] ASC
)WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON) ON [PRIMARY]
) ON [PRIMARY] TEXTIMAGE_ON [PRIMARY]
GO
ALTER TABLE [dbo].[NegociosFacturacionProveedorCredencial] ADD CONSTRAINT [DF_NegociosFacturacionProveedorCredencial_KeyVersion] DEFAULT (N'v1') FOR [KeyVersion]
GO
ALTER TABLE [dbo].[NegociosFacturacionProveedorCredencial] ADD CONSTRAINT [DF_NegociosFacturacionProveedorCredencial_Activo] DEFAULT ((1)) FOR [Activo]
GO
ALTER TABLE [dbo].[NegociosFacturacionProveedorCredencial] ADD CONSTRAINT [DF_NegociosFacturacionProveedorCredencial_FechaRegistro] DEFAULT (SYSUTCDATETIME()) FOR [FechaRegistro]
GO
ALTER TABLE [dbo].[NegociosFacturacionProveedorCredencial] ADD CONSTRAINT [CK_NegociosFacturacionProveedorCredencial_TipoCredencial]
CHECK ([TipoCredencial] IN (N'API_KEY', N'CLIENT_ID', N'CLIENT_SECRET', N'USUARIO', N'PASSWORD', N'TOKEN'))
GO
ALTER TABLE [dbo].[NegociosFacturacionProveedorCredencial]  WITH CHECK ADD CONSTRAINT [FK_NegociosFacturacionProveedorCredencial_NegociosFacturacionProveedorConfig_Id]
FOREIGN KEY([NegocioProveedorConfigId]) REFERENCES [dbo].[NegociosFacturacionProveedorConfig] ([Id])
GO
ALTER TABLE [dbo].[NegociosFacturacionProveedorCredencial] CHECK CONSTRAINT [FK_NegociosFacturacionProveedorCredencial_NegociosFacturacionProveedorConfig_Id]
GO
CREATE UNIQUE NONCLUSTERED INDEX [UX_NegociosFacturacionProveedorCredencial_Config_Tipo_Activo]
ON [dbo].[NegociosFacturacionProveedorCredencial]([NegocioProveedorConfigId] ASC, [TipoCredencial] ASC, [Activo] ASC)
GO

