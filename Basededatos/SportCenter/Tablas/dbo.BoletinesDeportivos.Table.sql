USE [DbSportCenter]
GO
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
-- =============================================
-- Author:        FRANCO LARA
-- Create date:   10/06/2026
-- Description:   Boletines/flyers de eventos deportivos cargados por usuarios publicos y super admin.
-- =============================================
CREATE TABLE [dbo].[BoletinesDeportivos](
    [IdBoletin] [int] IDENTITY(1,1) NOT NULL,
    [UsuarioId] [nvarchar](450) NOT NULL,
    [PerfilPublicoId] [int] NULL,
    [Titulo] [nvarchar](160) NULL,
    [Descripcion] [nvarchar](500) NULL,
    [ImagenUrl] [nvarchar](500) NOT NULL,
    [FechaEvento] [date] NOT NULL,
    [CodigoUbigeo] [char](6) NOT NULL,
    [TipoRegistro] [char](1) NOT NULL,
    [Activo] [bit] NOT NULL,
    [FechaCreacion] [datetime2](7) NOT NULL,
    [UsuarioCreacion] [nvarchar](120) NOT NULL,
    [FechaActualizacion] [datetime2](7) NULL,
    [UsuarioActualizacion] [nvarchar](120) NULL,
CONSTRAINT [PK_BoletinesDeportivos] PRIMARY KEY CLUSTERED ([IdBoletin] ASC)
) ON [PRIMARY];
GO

ALTER TABLE [dbo].[BoletinesDeportivos] ADD CONSTRAINT [DF_BoletinesDeportivos_TipoRegistro] DEFAULT ('U') FOR [TipoRegistro];
GO
ALTER TABLE [dbo].[BoletinesDeportivos] ADD CONSTRAINT [DF_BoletinesDeportivos_Activo] DEFAULT ((1)) FOR [Activo];
GO
ALTER TABLE [dbo].[BoletinesDeportivos] ADD CONSTRAINT [DF_BoletinesDeportivos_FechaCreacion] DEFAULT (SYSDATETIME()) FOR [FechaCreacion];
GO

ALTER TABLE [dbo].[BoletinesDeportivos] WITH CHECK ADD CONSTRAINT [CK_BoletinesDeportivos_TipoRegistro]
CHECK ([TipoRegistro] IN ('U', 'A'));
GO

ALTER TABLE [dbo].[BoletinesDeportivos] WITH CHECK ADD CONSTRAINT [FK_BoletinesDeportivos_AspNetUsers_UsuarioId]
FOREIGN KEY([UsuarioId]) REFERENCES [dbo].[AspNetUsers] ([Id]);
GO
ALTER TABLE [dbo].[BoletinesDeportivos] CHECK CONSTRAINT [FK_BoletinesDeportivos_AspNetUsers_UsuarioId];
GO

ALTER TABLE [dbo].[BoletinesDeportivos] WITH CHECK ADD CONSTRAINT [FK_BoletinesDeportivos_UsuariosPublicosPerfil_PerfilPublicoId]
FOREIGN KEY([PerfilPublicoId]) REFERENCES [dbo].[UsuariosPublicosPerfil] ([Id]);
GO
ALTER TABLE [dbo].[BoletinesDeportivos] CHECK CONSTRAINT [FK_BoletinesDeportivos_UsuariosPublicosPerfil_PerfilPublicoId];
GO

ALTER TABLE [dbo].[BoletinesDeportivos] WITH CHECK ADD CONSTRAINT [FK_BoletinesDeportivos_UbigeoDistritos_CodigoUbigeo]
FOREIGN KEY([CodigoUbigeo]) REFERENCES [dbo].[UbigeoDistritos] ([CodigoUbigeo]);
GO
ALTER TABLE [dbo].[BoletinesDeportivos] CHECK CONSTRAINT [FK_BoletinesDeportivos_UbigeoDistritos_CodigoUbigeo];
GO

CREATE NONCLUSTERED INDEX [IX_BoletinesDeportivos_Activo_FechaEvento]
ON [dbo].[BoletinesDeportivos] ([Activo], [FechaEvento] DESC, [FechaCreacion] DESC);
GO

CREATE NONCLUSTERED INDEX [IX_BoletinesDeportivos_UsuarioId_FechaCreacion]
ON [dbo].[BoletinesDeportivos] ([UsuarioId], [FechaCreacion] DESC);
GO

CREATE NONCLUSTERED INDEX [IX_BoletinesDeportivos_CodigoUbigeo_FechaEvento]
ON [dbo].[BoletinesDeportivos] ([CodigoUbigeo], [FechaEvento] DESC);
GO
