USE [DbSportCenter]
GO
/****** Object:  Table [dbo].[Negocios]    Script Date: 04/04/2026 ******/
-- Firma: Codex - 04/04/2026 | Agrega CodigoUbigeo en tabla Negocios y su relacion FK a UbigeoDistritos.
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE TABLE [dbo].[Negocios](
    [Id] [int] IDENTITY(1,1) NOT NULL,
    [NombreComercial] [nvarchar](max) NOT NULL,
    [RazonSocial] [nvarchar](max) NULL,
    [DocumentoFiscal] [nvarchar](max) NULL,
    [Activo] [bit] NOT NULL,
    [FechaRegistro] [datetime2](7) NOT NULL,
    [MonedaId] [int] NULL,
    [TipoDocumentoFiscal] [nvarchar](2) NULL,
    [NumeroDocumentoFiscal] [nvarchar](20) NULL,
    [DireccionFiscal] [nvarchar](250) NULL,
    [CodigoUbigeo] [char](6) NULL,
 CONSTRAINT [PK_Negocios] PRIMARY KEY CLUSTERED 
(
    [Id] ASC
)WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON, OPTIMIZE_FOR_SEQUENTIAL_KEY = OFF) ON [PRIMARY]
) ON [PRIMARY] TEXTIMAGE_ON [PRIMARY]
GO
ALTER TABLE [dbo].[Negocios] ADD  CONSTRAINT [DF_Negocios_MonedaId]  DEFAULT ((1)) FOR [MonedaId]
GO
ALTER TABLE [dbo].[Negocios]  WITH CHECK ADD  CONSTRAINT [FK_Negocios_Monedas_MonedaId] FOREIGN KEY([MonedaId])
REFERENCES [dbo].[Monedas] ([Id])
GO
ALTER TABLE [dbo].[Negocios] CHECK CONSTRAINT [FK_Negocios_Monedas_MonedaId]
GO
ALTER TABLE [dbo].[Negocios]  WITH CHECK ADD  CONSTRAINT [FK_Negocios_UbigeoDistritos_CodigoUbigeo] FOREIGN KEY([CodigoUbigeo])
REFERENCES [dbo].[UbigeoDistritos] ([CodigoUbigeo])
GO
ALTER TABLE [dbo].[Negocios] CHECK CONSTRAINT [FK_Negocios_UbigeoDistritos_CodigoUbigeo]
GO
