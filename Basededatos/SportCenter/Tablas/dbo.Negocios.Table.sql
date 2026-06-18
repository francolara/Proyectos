USE [DbSportCenter]
GO
/****** Object:  Table [dbo].[Negocios]    Script Date: 04/04/2026 ******/
-- Firma: Codex - 04/04/2026 | Agrega CodigoUbigeo en tabla Negocios y su relacion FK a UbigeoDistritos.
-- Firma: Codex - 06/04/2026 | Agrega politica de confirmacion de reserva por pago y porcentaje minimo de adelanto por negocio.
-- Firma: Codex - 09/04/2026 | Agrega configuracion de emision (CPE/Recibo interno) y porcentaje IGV.
-- Firma: Codex - 13/04/2026 | Agrega LogoUrl para imagen del logo del negocio.
-- Firma: Codex - 16/04/2026 | Agrega limites operativos (Sedes/Espacios) y flags de reserva (edicion de precio/cancelacion automatica por no confirmacion).
-- Firma: Codex - 19/04/2026 | Agrega UsuariosPermitidos como limite operativo adicional para gestion de usuarios por negocio.
-- Firma: FRANCO LARA - 18/06/2026 | Agrega TipoPlan en Negocios para distinguir capacidades Basico y Full.
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
    [PoliticaConfirmacionPago] [tinyint] NOT NULL,
    [PorcentajeAdelantoMinimo] [decimal](5,2) NULL,
    [PorcentajeIgv] [int] NOT NULL,
    [EmisionComprobantesElectronicos] [bit] NOT NULL,
    [EmisionReciboInterno] [bit] NOT NULL,
    [TipoDocumentoFiscal] [nvarchar](2) NULL,
    [NumeroDocumentoFiscal] [nvarchar](20) NULL,
    [DireccionFiscal] [nvarchar](250) NULL,
    [CodigoUbigeo] [char](6) NULL,
    [LogoUrl] [nvarchar](500) NULL,
    [PermitirModificarPrecioReserva] [bit] NOT NULL,
    [CancelacionAutomaticaNoConfirmada] [bit] NOT NULL,
    [MinutosCancelacionNoConfirmada] [int] NULL,
    [TipoPlan] [nvarchar](20) NOT NULL,
    [SedesPermitidas] [int] NOT NULL,
    [EspaciosPermitidos] [int] NOT NULL,
    [UsuariosPermitidos] [int] NOT NULL,
 CONSTRAINT [PK_Negocios] PRIMARY KEY CLUSTERED 
(
    [Id] ASC
)WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON, OPTIMIZE_FOR_SEQUENTIAL_KEY = OFF) ON [PRIMARY]
) ON [PRIMARY] TEXTIMAGE_ON [PRIMARY]
GO
ALTER TABLE [dbo].[Negocios] ADD  CONSTRAINT [DF_Negocios_MonedaId]  DEFAULT ((1)) FOR [MonedaId]
GO
ALTER TABLE [dbo].[Negocios] ADD  CONSTRAINT [DF_Negocios_PoliticaConfirmacionPago]  DEFAULT ((0)) FOR [PoliticaConfirmacionPago]
GO
ALTER TABLE [dbo].[Negocios] ADD  CONSTRAINT [DF_Negocios_PorcentajeIgv]  DEFAULT ((18)) FOR [PorcentajeIgv]
GO
ALTER TABLE [dbo].[Negocios] ADD  CONSTRAINT [DF_Negocios_EmisionComprobantesElectronicos]  DEFAULT ((0)) FOR [EmisionComprobantesElectronicos]
GO
ALTER TABLE [dbo].[Negocios] ADD  CONSTRAINT [DF_Negocios_EmisionReciboInterno]  DEFAULT ((0)) FOR [EmisionReciboInterno]
GO
ALTER TABLE [dbo].[Negocios] ADD  CONSTRAINT [DF_Negocios_PermitirModificarPrecioReserva]  DEFAULT ((0)) FOR [PermitirModificarPrecioReserva]
GO
ALTER TABLE [dbo].[Negocios] ADD  CONSTRAINT [DF_Negocios_CancelacionAutomaticaNoConfirmada]  DEFAULT ((0)) FOR [CancelacionAutomaticaNoConfirmada]
GO
ALTER TABLE [dbo].[Negocios] ADD  CONSTRAINT [DF_Negocios_TipoPlan]  DEFAULT (N'Basico') FOR [TipoPlan]
GO
ALTER TABLE [dbo].[Negocios] ADD  CONSTRAINT [DF_Negocios_SedesPermitidas]  DEFAULT ((2)) FOR [SedesPermitidas]
GO
ALTER TABLE [dbo].[Negocios] ADD  CONSTRAINT [DF_Negocios_EspaciosPermitidos]  DEFAULT ((6)) FOR [EspaciosPermitidos]
GO
ALTER TABLE [dbo].[Negocios] ADD  CONSTRAINT [DF_Negocios_UsuariosPermitidos]  DEFAULT ((3)) FOR [UsuariosPermitidos]
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
