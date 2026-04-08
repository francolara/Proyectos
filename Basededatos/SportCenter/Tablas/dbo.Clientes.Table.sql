USE [DbSportCenter]
GO
/****** Object:  Table [dbo].[Clientes]    Script Date: 04/04/2026 ******/
-- Firma: Codex - 04/04/2026 | Agrega CodigoUbigeo en tabla Clientes y su relacion FK a UbigeoDistritos.
-- Firma: Codex - 06/04/2026 | Agrega columnas Nombres y Apellidos para clientes naturales, manteniendo NombresORazonSocial para compatibilidad de listados.
-- Firma: Codex - 06/04/2026 | Cliente queda asociado directamente al NegocioId; se elimina tabla puente NegocioClientes.
-- Firma: Codex - 07/04/2026 | Indice unico excluye tipo documento 0 para permitir multiples clientes no domiciliados sin RUC.
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE TABLE [dbo].[Clientes](
    [Id] [int] IDENTITY(1,1) NOT NULL,
    [NegocioId] [int] NOT NULL,
    [NombresORazonSocial] [nvarchar](200) NOT NULL,
    [Nombres] [nvarchar](120) NULL,
    [Apellidos] [nvarchar](120) NULL,
    [TipoDocumento] [nvarchar](2) NOT NULL,
    [NumeroDocumento] [nvarchar](20) NOT NULL,
    [Telefono] [nvarchar](20) NULL,
    [Correo] [nvarchar](200) NULL,
    [Activo] [bit] NOT NULL,
    [DireccionFiscal] [nvarchar](250) NULL,
    [FechaCreacion] [datetime2](7) NOT NULL,
    [UsuarioCreacion] [nvarchar](200) NULL,
    [FechaActualizacion] [datetime2](7) NULL,
    [UsuarioActualizacion] [nvarchar](200) NULL,
    [NombreEquipo] [nvarchar](120) NULL,
    [CodigoUbigeo] [char](6) NULL,
 CONSTRAINT [PK_Clientes] PRIMARY KEY CLUSTERED 
(
    [Id] ASC
)WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON, OPTIMIZE_FOR_SEQUENTIAL_KEY = OFF) ON [PRIMARY]
) ON [PRIMARY]
GO
ALTER TABLE [dbo].[Clientes] ADD  CONSTRAINT [DF_Clientes_FechaCreacion]  DEFAULT (sysutcdatetime()) FOR [FechaCreacion]
GO
ALTER TABLE [dbo].[Clientes]  WITH CHECK ADD  CONSTRAINT [FK_Clientes_UbigeoDistritos_CodigoUbigeo] FOREIGN KEY([CodigoUbigeo])
REFERENCES [dbo].[UbigeoDistritos] ([CodigoUbigeo])
GO
ALTER TABLE [dbo].[Clientes] CHECK CONSTRAINT [FK_Clientes_UbigeoDistritos_CodigoUbigeo]
GO
ALTER TABLE [dbo].[Clientes]  WITH CHECK ADD  CONSTRAINT [FK_Clientes_Negocios_NegocioId] FOREIGN KEY([NegocioId])
REFERENCES [dbo].[Negocios] ([Id])
GO
ALTER TABLE [dbo].[Clientes] CHECK CONSTRAINT [FK_Clientes_Negocios_NegocioId]
GO
CREATE UNIQUE NONCLUSTERED INDEX [UX_Clientes_Negocio_Tipo_Numero_Activo]
ON [dbo].[Clientes] ([NegocioId] ASC, [TipoDocumento] ASC, [NumeroDocumento] ASC)
WHERE [Activo] = 1 AND [TipoDocumento] <> N'0'
GO
