USE [DbSportCenter]
GO
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
-- =============================================
-- Author:        FRANCO LARA
-- Create date:   16/04/2026
-- Description:   Relacion entre reserva creada por portal y usuario publico autenticado.
-- =============================================
CREATE TABLE [dbo].[ReservasUsuariosPublicos](
    [ReservaId] [int] NOT NULL,
    [UsuarioId] [nvarchar](450) NOT NULL,
    [FechaCreacion] [datetime2](7) NOT NULL,
    [UsuarioCreacion] [nvarchar](120) NOT NULL,
CONSTRAINT [PK_ReservasUsuariosPublicos] PRIMARY KEY CLUSTERED ([ReservaId] ASC, [UsuarioId] ASC)
) ON [PRIMARY];
GO

ALTER TABLE [dbo].[ReservasUsuariosPublicos] ADD CONSTRAINT [DF_ReservasUsuariosPublicos_FechaCreacion] DEFAULT (SYSDATETIME()) FOR [FechaCreacion];
GO

ALTER TABLE [dbo].[ReservasUsuariosPublicos] WITH CHECK ADD CONSTRAINT [FK_ReservasUsuariosPublicos_Reservas_ReservaId]
FOREIGN KEY([ReservaId]) REFERENCES [dbo].[Reservas] ([Id]);
GO
ALTER TABLE [dbo].[ReservasUsuariosPublicos] CHECK CONSTRAINT [FK_ReservasUsuariosPublicos_Reservas_ReservaId];
GO
ALTER TABLE [dbo].[ReservasUsuariosPublicos] WITH CHECK ADD CONSTRAINT [FK_ReservasUsuariosPublicos_AspNetUsers_UsuarioId]
FOREIGN KEY([UsuarioId]) REFERENCES [dbo].[AspNetUsers] ([Id]);
GO
ALTER TABLE [dbo].[ReservasUsuariosPublicos] CHECK CONSTRAINT [FK_ReservasUsuariosPublicos_AspNetUsers_UsuarioId];
GO
