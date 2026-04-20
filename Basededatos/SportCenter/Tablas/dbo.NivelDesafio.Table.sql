USE [DbSportCenter]
GO
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
-- =============================================
-- Author:        FRANCO LARA
-- Create date:   18/04/2026
-- Description:   Catalogo de niveles para desafios entre equipos.
-- =============================================
CREATE TABLE [dbo].[NivelDesafio](
    [IdNivel] [int] IDENTITY(1,1) NOT NULL,
    [Nombre] [nvarchar](80) NOT NULL,
    [Activo] [bit] NOT NULL,
    [Orden] [int] NOT NULL,
CONSTRAINT [PK_NivelDesafio] PRIMARY KEY CLUSTERED ([IdNivel] ASC)
) ON [PRIMARY];
GO
ALTER TABLE [dbo].[NivelDesafio] ADD CONSTRAINT [DF_NivelDesafio_Activo] DEFAULT ((1)) FOR [Activo];
GO
CREATE UNIQUE NONCLUSTERED INDEX [UX_NivelDesafio_Nombre]
    ON [dbo].[NivelDesafio]([Nombre] ASC);
GO
