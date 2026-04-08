USE [DbSportCenter]
GO
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
-- =============================================
-- Author:        FRANCO LARA
-- Create date:   06/04/2026
-- Description:   Tabla supermaestro para tipos de deporte.
-- =============================================
CREATE TABLE [dbo].[TiposDeporteSuperMaestro](
    [Id] [int] IDENTITY(1,1) NOT NULL,
    [Codigo] [nvarchar](20) NOT NULL,
    [Nombre] [nvarchar](120) NOT NULL,
    [Activo] [bit] NOT NULL,
    [FechaCreacion] [datetime2](7) NOT NULL,
    [UsuarioCreacion] [nvarchar](200) NULL,
    [FechaActualizacion] [datetime2](7) NULL,
    [UsuarioActualizacion] [nvarchar](200) NULL,
 CONSTRAINT [PK_TiposDeporteSuperMaestro] PRIMARY KEY CLUSTERED
(
    [Id] ASC
) WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON)
) ON [PRIMARY]
GO
ALTER TABLE [dbo].[TiposDeporteSuperMaestro] ADD CONSTRAINT [DF_TiposDeporteSuperMaestro_Activo] DEFAULT ((1)) FOR [Activo]
GO
ALTER TABLE [dbo].[TiposDeporteSuperMaestro] ADD CONSTRAINT [DF_TiposDeporteSuperMaestro_FechaCreacion] DEFAULT (sysutcdatetime()) FOR [FechaCreacion]
GO
CREATE UNIQUE NONCLUSTERED INDEX [UQ_TiposDeporteSuperMaestro_Codigo]
    ON [dbo].[TiposDeporteSuperMaestro]([Codigo] ASC)
GO
CREATE UNIQUE NONCLUSTERED INDEX [UQ_TiposDeporteSuperMaestro_Nombre]
    ON [dbo].[TiposDeporteSuperMaestro]([Nombre] ASC)
GO