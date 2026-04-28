USE [DbSportCenter]
GO
/****** Object:  Table [dbo].[Sedes]    Script Date: 3/04/2026 23:17:42 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
-- =============================================
-- Author:        FRANCO LARA
-- Create date:   06/04/2026
-- Description:   Agrega campo ConsideracionesReserva para publicar reglas y condiciones de reserva por sede.
-- =============================================
-- =============================================
-- Author:        FRANCO LARA
-- Create date:   16/04/2026
-- Description:   Agrega URLs de redes sociales (Facebook/Instagram/Twitter) por sede para portal publico.
-- =============================================
-- =============================================
-- Author:        FRANCO LARA
-- Create date:   27/04/2026
-- Description:   Agrega CodigoUbigeo por sede para filtros publicos por ubicacion real de cada sede.
-- =============================================
CREATE TABLE [dbo].[Sedes](
	[Id] [int] IDENTITY(1,1) NOT NULL,
	[NegocioId] [int] NOT NULL,
	[Nombre] [nvarchar](max) NOT NULL,
	[Direccion] [nvarchar](max) NOT NULL,
	[CodigoUbigeo] [char](6) NULL,
	[ConsideracionesReserva] [nvarchar](2000) NULL,
	[Telefono] [nvarchar](max) NULL,
	[FacebookUrl] [nvarchar](500) NULL,
	[InstagramUrl] [nvarchar](500) NULL,
	[TwitterUrl] [nvarchar](500) NULL,
	[Activo] [bit] NOT NULL,
	[FechaActualizacion] [datetime2](7) NULL,
	[FechaCreacion] [datetime2](7) NOT NULL,
	[UsuarioActualizacion] [nvarchar](max) NULL,
	[UsuarioCreacion] [nvarchar](max) NULL,
	[Latitud] [decimal](10, 7) NULL,
	[Longitud] [decimal](10, 7) NULL,
	[GooglePlaceId] [nvarchar](200) NULL,
	[GoogleMapsUrl] [nvarchar](500) NULL,
	[FotoPrincipalUrl] [nvarchar](500) NULL,
	[FotosUrlsCsv] [nvarchar](max) NULL,
 CONSTRAINT [PK_Sedes] PRIMARY KEY CLUSTERED 
(
	[Id] ASC
)WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON, OPTIMIZE_FOR_SEQUENTIAL_KEY = OFF) ON [PRIMARY]
) ON [PRIMARY] TEXTIMAGE_ON [PRIMARY]
GO
ALTER TABLE [dbo].[Sedes] ADD  DEFAULT ('0001-01-01T00:00:00.0000000') FOR [FechaCreacion]
GO
ALTER TABLE [dbo].[Sedes]  WITH CHECK ADD  CONSTRAINT [FK_Sedes_Negocios_NegocioId] FOREIGN KEY([NegocioId])
REFERENCES [dbo].[Negocios] ([Id])
ON DELETE CASCADE
GO
ALTER TABLE [dbo].[Sedes] CHECK CONSTRAINT [FK_Sedes_Negocios_NegocioId]
GO
ALTER TABLE [dbo].[Sedes]  WITH CHECK ADD  CONSTRAINT [FK_Sedes_UbigeoDistritos_CodigoUbigeo] FOREIGN KEY([CodigoUbigeo])
REFERENCES [dbo].[UbigeoDistritos] ([CodigoUbigeo])
GO
ALTER TABLE [dbo].[Sedes] CHECK CONSTRAINT [FK_Sedes_UbigeoDistritos_CodigoUbigeo]
GO
