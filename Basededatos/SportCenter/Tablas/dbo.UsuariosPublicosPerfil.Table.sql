USE [DbSportCenter]
GO
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
-- =============================================
-- Author:        FRANCO LARA
-- Create date:   16/04/2026
-- Description:   Perfil de usuario publico para portal (datos personales, documento, ubigeo, equipo y fecha de nacimiento).
-- =============================================
-- =============================================
-- Author:        FRANCO LARA
-- Create date:   18/04/2026
-- Description:   Agrega configuracion de desafios y detalle general del equipo al perfil publico.
-- =============================================
-- =============================================
-- Author:        FRANCO LARA
-- Create date:   18/04/2026
-- Description:   Registra ubicacion propia y WhatsApp operativo del equipo para mostrar perfil y coordinar desafios.
-- =============================================
CREATE TABLE [dbo].[UsuariosPublicosPerfil](
    [Id] [int] IDENTITY(1,1) NOT NULL,
    [UsuarioId] [nvarchar](450) NOT NULL,
    [TipoDocumento] [nvarchar](20) NOT NULL,
    [NumeroDocumento] [nvarchar](20) NULL,
    [Nombres] [nvarchar](120) NOT NULL,
    [Apellidos] [nvarchar](120) NOT NULL,
    [NombreEquipo] [nvarchar](120) NULL,
    [Telefono] [nvarchar](30) NULL,
    [Correo] [nvarchar](200) NULL,
    [FechaNacimiento] [date] NULL,
    [CodigoUbigeo] [char](6) NULL,
    [BuscarDesafios] [bit] NOT NULL,
    [IdDeporteDesafio] [int] NULL,
    [IdNivelDesafio] [int] NULL,
    [ObservacionDesafio] [nvarchar](500) NULL,
    [DetalleEquipo] [nvarchar](1000) NULL,
    [CodigoUbigeoEquipo] [char](6) NULL,
    [WhatsappEquipo] [nvarchar](30) NULL,
    [FechaCreacion] [datetime2](7) NOT NULL,
    [UsuarioCreacion] [nvarchar](120) NOT NULL,
    [FechaActualizacion] [datetime2](7) NULL,
    [UsuarioActualizacion] [nvarchar](120) NULL,
CONSTRAINT [PK_UsuariosPublicosPerfil] PRIMARY KEY CLUSTERED ([Id] ASC),
CONSTRAINT [UQ_UsuariosPublicosPerfil_UsuarioId] UNIQUE NONCLUSTERED ([UsuarioId] ASC)
) ON [PRIMARY];
GO

ALTER TABLE [dbo].[UsuariosPublicosPerfil] ADD CONSTRAINT [DF_UsuariosPublicosPerfil_TipoDocumento] DEFAULT (N'0') FOR [TipoDocumento];
GO
ALTER TABLE [dbo].[UsuariosPublicosPerfil] ADD CONSTRAINT [DF_UsuariosPublicosPerfil_BuscarDesafios] DEFAULT ((0)) FOR [BuscarDesafios];
GO
ALTER TABLE [dbo].[UsuariosPublicosPerfil] ADD CONSTRAINT [DF_UsuariosPublicosPerfil_FechaCreacion] DEFAULT (SYSDATETIME()) FOR [FechaCreacion];
GO

ALTER TABLE [dbo].[UsuariosPublicosPerfil]  WITH CHECK ADD CONSTRAINT [FK_UsuariosPublicosPerfil_AspNetUsers_UsuarioId]
FOREIGN KEY([UsuarioId]) REFERENCES [dbo].[AspNetUsers] ([Id]);
GO
ALTER TABLE [dbo].[UsuariosPublicosPerfil] CHECK CONSTRAINT [FK_UsuariosPublicosPerfil_AspNetUsers_UsuarioId];
GO
ALTER TABLE [dbo].[UsuariosPublicosPerfil]  WITH CHECK ADD CONSTRAINT [FK_UsuariosPublicosPerfil_UbigeoDistritos_CodigoUbigeo]
FOREIGN KEY([CodigoUbigeo]) REFERENCES [dbo].[UbigeoDistritos] ([CodigoUbigeo]);
GO
ALTER TABLE [dbo].[UsuariosPublicosPerfil] CHECK CONSTRAINT [FK_UsuariosPublicosPerfil_UbigeoDistritos_CodigoUbigeo];
GO
ALTER TABLE [dbo].[UsuariosPublicosPerfil]  WITH CHECK ADD CONSTRAINT [FK_UsuariosPublicosPerfil_TiposDeporte_IdDeporteDesafio]
FOREIGN KEY([IdDeporteDesafio]) REFERENCES [dbo].[TiposDeporte] ([Id]);
GO
ALTER TABLE [dbo].[UsuariosPublicosPerfil] CHECK CONSTRAINT [FK_UsuariosPublicosPerfil_TiposDeporte_IdDeporteDesafio];
GO
ALTER TABLE [dbo].[UsuariosPublicosPerfil]  WITH CHECK ADD CONSTRAINT [FK_UsuariosPublicosPerfil_NivelDesafio_IdNivelDesafio]
FOREIGN KEY([IdNivelDesafio]) REFERENCES [dbo].[NivelDesafio] ([IdNivel]);
GO
ALTER TABLE [dbo].[UsuariosPublicosPerfil] CHECK CONSTRAINT [FK_UsuariosPublicosPerfil_NivelDesafio_IdNivelDesafio];
GO
ALTER TABLE [dbo].[UsuariosPublicosPerfil]  WITH CHECK ADD CONSTRAINT [FK_UsuariosPublicosPerfil_UbigeoDistritos_CodigoUbigeoEquipo]
FOREIGN KEY([CodigoUbigeoEquipo]) REFERENCES [dbo].[UbigeoDistritos] ([CodigoUbigeo]);
GO
ALTER TABLE [dbo].[UsuariosPublicosPerfil] CHECK CONSTRAINT [FK_UsuariosPublicosPerfil_UbigeoDistritos_CodigoUbigeoEquipo];
GO
