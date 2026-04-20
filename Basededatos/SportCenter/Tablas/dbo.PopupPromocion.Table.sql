USE [DbSportCenter]
GO
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO

-- Firma: Codex - 20/04/2026 | Tabla global de anuncios popup para home publico, administrable desde super admin, con orientacion por anuncio para mezclar piezas verticales y horizontales, y subtitulo opcional por pieza.
CREATE TABLE [dbo].[PopupPromocion](
    [IdPopupPromocion] [int] IDENTITY(1,1) NOT NULL,
    [Titulo] [nvarchar](120) NOT NULL,
    [Subtitulo] [nvarchar](140) NULL,
    [Descripcion] [nvarchar](260) NULL,
    [ImagenUrl] [nvarchar](500) NOT NULL,
    [Orientacion] [char](1) NOT NULL,
    [TextoBoton] [nvarchar](40) NULL,
    [UrlBoton] [nvarchar](300) NULL,
    [UrlImagen] [nvarchar](300) NULL,
    [Orden] [int] NOT NULL,
    [Activo] [bit] NOT NULL,
    [FechaInicio] [date] NULL,
    [FechaFin] [date] NULL,
    [AbrirNuevaPestana] [bit] NOT NULL,
    [FechaCreacion] [datetime2](7) NOT NULL,
    [FechaModificacion] [datetime2](7) NULL,
 CONSTRAINT [PK_PopupPromocion] PRIMARY KEY CLUSTERED
(
    [IdPopupPromocion] ASC
)
) ON [PRIMARY]
GO

ALTER TABLE [dbo].[PopupPromocion]
ADD CONSTRAINT [DF_PopupPromocion_Orden] DEFAULT ((1)) FOR [Orden]
GO

ALTER TABLE [dbo].[PopupPromocion]
ADD CONSTRAINT [DF_PopupPromocion_Activo] DEFAULT ((1)) FOR [Activo]
GO

ALTER TABLE [dbo].[PopupPromocion]
ADD CONSTRAINT [DF_PopupPromocion_Orientacion] DEFAULT ('V') FOR [Orientacion]
GO

ALTER TABLE [dbo].[PopupPromocion]
ADD CONSTRAINT [DF_PopupPromocion_AbrirNuevaPestana] DEFAULT ((1)) FOR [AbrirNuevaPestana]
GO

ALTER TABLE [dbo].[PopupPromocion]
ADD CONSTRAINT [DF_PopupPromocion_FechaCreacion] DEFAULT (SYSDATETIME()) FOR [FechaCreacion]
GO

ALTER TABLE [dbo].[PopupPromocion]
ADD CONSTRAINT [CK_PopupPromocion_Orientacion] CHECK ([Orientacion] IN ('V','H'))
GO

CREATE INDEX [IX_PopupPromocion_Activo_Fechas_Orden]
ON [dbo].[PopupPromocion] ([Activo], [FechaInicio], [FechaFin], [Orden], [FechaCreacion])
GO
