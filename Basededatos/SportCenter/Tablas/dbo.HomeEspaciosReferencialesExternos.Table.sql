USE [DbSportCenter]
GO
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
-- =============================================
-- Author:        FRANCO LARA
-- Create date:   27/04/2026
-- Description:   Tabla adicional para listar espacios/complejos referenciales externos en Home (no afiliados), usados en union del buscador publico.
-- Firma: Codex - 27/04/2026 | Se agrega TelefonoContacto para almacenar el numero obtenido desde Google Places Details.
-- =============================================
CREATE TABLE [dbo].[HomeEspaciosReferencialesExternos](
    [Id] [int] IDENTITY(1,1) NOT NULL,
    [GooglePlaceId] [nvarchar](200) NULL,
    [NombreComplejo] [nvarchar](180) NOT NULL,
    [NombreEspacio] [nvarchar](150) NULL,
    [CodigoReferencia] [nvarchar](50) NULL,
    [CodigoUbigeo] [char](6) NOT NULL,
    [TipoDeporteSuperId] [int] NOT NULL,
    [Direccion] [nvarchar](250) NULL,
    [Referencia] [nvarchar](1000) NULL,
    [TelefonoContacto] [nvarchar](40) NULL,
    [CorreoContacto] [nvarchar](200) NULL,
    [WhatsappContacto] [nvarchar](30) NULL,
    [PermiteChatWhatsapp] [bit] NOT NULL,
    [TarifaReferencial] [decimal](10,2) NULL,
    [TieneIluminacion] [bit] NOT NULL,
    [Techada] [bit] NOT NULL,
    [GoogleMapsUrl] [nvarchar](500) NULL,
    [LatitudReferencia] [decimal](10,7) NULL,
    [LongitudReferencia] [decimal](10,7) NULL,
    [FotoPrincipalUrl] [nvarchar](500) NULL,
    [FotosUrlsCsv] [nvarchar](max) NULL,
    [Activo] [bit] NOT NULL,
    [FechaCreacion] [datetime2](7) NOT NULL,
    [UsuarioCreacion] [nvarchar](200) NULL,
    [FechaActualizacion] [datetime2](7) NULL,
    [UsuarioActualizacion] [nvarchar](200) NULL,
 CONSTRAINT [PK_HomeEspaciosReferencialesExternos] PRIMARY KEY CLUSTERED
(
    [Id] ASC
) WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON)
) ON [PRIMARY] TEXTIMAGE_ON [PRIMARY]
GO
ALTER TABLE [dbo].[HomeEspaciosReferencialesExternos] ADD CONSTRAINT [DF_HomeEspaciosReferencialesExternos_Activo] DEFAULT ((1)) FOR [Activo]
GO
ALTER TABLE [dbo].[HomeEspaciosReferencialesExternos] ADD CONSTRAINT [DF_HomeEspaciosReferencialesExternos_PermiteChatWhatsapp] DEFAULT ((1)) FOR [PermiteChatWhatsapp]
GO
ALTER TABLE [dbo].[HomeEspaciosReferencialesExternos] ADD CONSTRAINT [DF_HomeEspaciosReferencialesExternos_TieneIluminacion] DEFAULT ((0)) FOR [TieneIluminacion]
GO
ALTER TABLE [dbo].[HomeEspaciosReferencialesExternos] ADD CONSTRAINT [DF_HomeEspaciosReferencialesExternos_Techada] DEFAULT ((0)) FOR [Techada]
GO
ALTER TABLE [dbo].[HomeEspaciosReferencialesExternos] ADD CONSTRAINT [DF_HomeEspaciosReferencialesExternos_FechaCreacion] DEFAULT (sysutcdatetime()) FOR [FechaCreacion]
GO
ALTER TABLE [dbo].[HomeEspaciosReferencialesExternos]  WITH CHECK ADD  CONSTRAINT [FK_HomeEspaciosReferencialesExternos_UbigeoDistritos_CodigoUbigeo] FOREIGN KEY([CodigoUbigeo])
REFERENCES [dbo].[UbigeoDistritos] ([CodigoUbigeo])
GO
ALTER TABLE [dbo].[HomeEspaciosReferencialesExternos] CHECK CONSTRAINT [FK_HomeEspaciosReferencialesExternos_UbigeoDistritos_CodigoUbigeo]
GO
ALTER TABLE [dbo].[HomeEspaciosReferencialesExternos]  WITH CHECK ADD  CONSTRAINT [FK_HomeEspaciosReferencialesExternos_TiposDeporteSuperMaestro_TipoDeporteSuperId] FOREIGN KEY([TipoDeporteSuperId])
REFERENCES [dbo].[TiposDeporteSuperMaestro] ([Id])
GO
ALTER TABLE [dbo].[HomeEspaciosReferencialesExternos] CHECK CONSTRAINT [FK_HomeEspaciosReferencialesExternos_TiposDeporteSuperMaestro_TipoDeporteSuperId]
GO
CREATE NONCLUSTERED INDEX [IX_HomeEspaciosReferencialesExternos_Busqueda]
    ON [dbo].[HomeEspaciosReferencialesExternos]([Activo] ASC, [TipoDeporteSuperId] ASC, [CodigoUbigeo] ASC)
GO
CREATE UNIQUE NONCLUSTERED INDEX [UQ_HomeEspaciosReferencialesExternos_GooglePlaceId]
    ON [dbo].[HomeEspaciosReferencialesExternos]([GooglePlaceId] ASC)
    WHERE [GooglePlaceId] IS NOT NULL
GO
