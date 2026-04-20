USE [DbSportCenter]
GO
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
-- =============================================
-- Author:        FRANCO LARA
-- Create date:   18/04/2026
-- Description:   Tabla de mensajeria interna para coordinacion de desafios entre equipos.
-- =============================================
CREATE TABLE [dbo].[DesafioMensaje](
    [IdMensaje] [int] IDENTITY(1,1) NOT NULL,
    [IdDesafio] [int] NOT NULL,
    [UsuarioIdEmisor] [nvarchar](450) NOT NULL,
    [Mensaje] [nvarchar](500) NOT NULL,
    [FechaRegistro] [datetime2](7) NOT NULL,
    [Activo] [bit] NOT NULL,
    [UsuarioCreacion] [nvarchar](120) NOT NULL,
CONSTRAINT [PK_DesafioMensaje] PRIMARY KEY CLUSTERED ([IdMensaje] ASC)
) ON [PRIMARY];
GO
ALTER TABLE [dbo].[DesafioMensaje] ADD CONSTRAINT [DF_DesafioMensaje_FechaRegistro] DEFAULT (SYSDATETIME()) FOR [FechaRegistro];
GO
ALTER TABLE [dbo].[DesafioMensaje] ADD CONSTRAINT [DF_DesafioMensaje_Activo] DEFAULT ((1)) FOR [Activo];
GO
ALTER TABLE [dbo].[DesafioMensaje]  WITH CHECK ADD CONSTRAINT [FK_DesafioMensaje_Desafio_IdDesafio]
FOREIGN KEY([IdDesafio]) REFERENCES [dbo].[Desafio] ([Id]);
GO
ALTER TABLE [dbo].[DesafioMensaje] CHECK CONSTRAINT [FK_DesafioMensaje_Desafio_IdDesafio];
GO
ALTER TABLE [dbo].[DesafioMensaje]  WITH CHECK ADD CONSTRAINT [FK_DesafioMensaje_AspNetUsers_UsuarioIdEmisor]
FOREIGN KEY([UsuarioIdEmisor]) REFERENCES [dbo].[AspNetUsers] ([Id]);
GO
ALTER TABLE [dbo].[DesafioMensaje] CHECK CONSTRAINT [FK_DesafioMensaje_AspNetUsers_UsuarioIdEmisor];
GO
CREATE NONCLUSTERED INDEX [IX_DesafioMensaje_Desafio_FechaRegistro]
    ON [dbo].[DesafioMensaje]([IdDesafio] ASC, [FechaRegistro] ASC);
GO
