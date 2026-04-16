USE [DbSportCenter]
GO
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO

-- Firma: Codex - 14/04/2026 | Tabla global de banners web (Home/Login/Registro) administrable desde panel.
CREATE TABLE [dbo].[WebBanners](
    [Id] [int] IDENTITY(1,1) NOT NULL,
    [Titulo] [nvarchar](120) NOT NULL,
    [Subtitulo] [nvarchar](220) NULL,
    [Descripcion] [nvarchar](400) NULL,
    [BotonTexto] [nvarchar](40) NULL,
    [BotonUrl] [nvarchar](300) NULL,
    [ImagenUrl] [nvarchar](500) NOT NULL,
    [ImagenUrlMobile] [nvarchar](500) NULL,
    [TipoBanner] [tinyint] NOT NULL,
    [Orden] [int] NOT NULL,
    [Activo] [bit] NOT NULL,
    [FechaInicio] [date] NULL,
    [FechaFin] [date] NULL,
    [FechaRegistro] [datetime2](7) NOT NULL,
    [FechaActualizacion] [datetime2](7) NULL,
    [UsuarioRegistro] [nvarchar](120) NULL,
    [UsuarioActualizacion] [nvarchar](120) NULL,
 CONSTRAINT [PK_WebBanners] PRIMARY KEY CLUSTERED
(
    [Id] ASC
)
) ON [PRIMARY]
GO

ALTER TABLE [dbo].[WebBanners]
ADD CONSTRAINT [DF_WebBanners_Orden] DEFAULT ((1)) FOR [Orden]
GO

ALTER TABLE [dbo].[WebBanners]
ADD CONSTRAINT [DF_WebBanners_TipoBanner] DEFAULT ((1)) FOR [TipoBanner]
GO

ALTER TABLE [dbo].[WebBanners]
ADD CONSTRAINT [DF_WebBanners_Activo] DEFAULT ((1)) FOR [Activo]
GO

ALTER TABLE [dbo].[WebBanners]
ADD CONSTRAINT [DF_WebBanners_FechaRegistro] DEFAULT (SYSDATETIME()) FOR [FechaRegistro]
GO

CREATE INDEX [IX_WebBanners_Activo_Orden]
ON [dbo].[WebBanners] ([Activo], [Orden], [Id])
GO
