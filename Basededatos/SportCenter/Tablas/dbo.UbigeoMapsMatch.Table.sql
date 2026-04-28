USE [DbSportCenter]
GO
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
-- =============================================
-- Author:        FRANCO LARA
-- Create date:   27/04/2026
-- Description:   Diccionario de equivalencias Google Maps -> UBIGEO SUNAT para resolver ubicacion de sedes por PlaceId y/o nombres de componentes.
-- Firma:         Codex - 27/04/2026 | Tabla base para persistir match de ubicaciones Google con maestro SUNAT.
-- =============================================
CREATE TABLE [dbo].[UbigeoMapsMatch](
	[Id] [int] IDENTITY(1,1) NOT NULL,
	[CountryCode] [char](2) NOT NULL CONSTRAINT [DF_UbigeoMapsMatch_CountryCode] DEFAULT ('PE'),
	[GooglePlaceId] [nvarchar](200) NULL,
	[GoogleDepartamento] [nvarchar](120) NULL,
	[GoogleProvincia] [nvarchar](120) NULL,
	[GoogleDistrito] [nvarchar](120) NULL,
	[CodigoUbigeo] [char](6) NOT NULL,
	[EsManual] [bit] NOT NULL CONSTRAINT [DF_UbigeoMapsMatch_EsManual] DEFAULT ((0)),
	[Activo] [bit] NOT NULL CONSTRAINT [DF_UbigeoMapsMatch_Activo] DEFAULT ((1)),
	[FechaCreacion] [datetime2](7) NOT NULL CONSTRAINT [DF_UbigeoMapsMatch_FechaCreacion] DEFAULT (sysutcdatetime()),
	[UsuarioCreacion] [nvarchar](200) NULL,
	[FechaActualizacion] [datetime2](7) NULL,
	[UsuarioActualizacion] [nvarchar](200) NULL,
	[GoogleDepartamentoNorm] AS UPPER(LTRIM(RTRIM([GoogleDepartamento]))) PERSISTED,
	[GoogleProvinciaNorm] AS UPPER(LTRIM(RTRIM([GoogleProvincia]))) PERSISTED,
	[GoogleDistritoNorm] AS UPPER(LTRIM(RTRIM([GoogleDistrito]))) PERSISTED,
 CONSTRAINT [PK_UbigeoMapsMatch] PRIMARY KEY CLUSTERED
(
	[Id] ASC
))
GO
ALTER TABLE [dbo].[UbigeoMapsMatch]  WITH CHECK ADD  CONSTRAINT [FK_UbigeoMapsMatch_UbigeoDistritos_CodigoUbigeo] FOREIGN KEY([CodigoUbigeo])
REFERENCES [dbo].[UbigeoDistritos] ([CodigoUbigeo])
GO
ALTER TABLE [dbo].[UbigeoMapsMatch] CHECK CONSTRAINT [FK_UbigeoMapsMatch_UbigeoDistritos_CodigoUbigeo]
GO
CREATE UNIQUE NONCLUSTERED INDEX [UX_UbigeoMapsMatch_GooglePlaceId]
ON [dbo].[UbigeoMapsMatch]([GooglePlaceId] ASC)
WHERE ([GooglePlaceId] IS NOT NULL)
GO
CREATE UNIQUE NONCLUSTERED INDEX [UX_UbigeoMapsMatch_Texto]
ON [dbo].[UbigeoMapsMatch]([CountryCode] ASC, [GoogleDepartamentoNorm] ASC, [GoogleProvinciaNorm] ASC, [GoogleDistritoNorm] ASC)
WHERE ([GoogleDepartamentoNorm] IS NOT NULL AND [GoogleProvinciaNorm] IS NOT NULL AND [GoogleDistritoNorm] IS NOT NULL)
GO
