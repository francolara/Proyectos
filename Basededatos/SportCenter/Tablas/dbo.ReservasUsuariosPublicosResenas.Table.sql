
GO
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
-- =============================================
-- Author:        FRANCO LARA
-- Create date:   08/06/2026
-- Description:   Almacena una sola resena publica por reserva creada por usuario publico autenticado.
-- =============================================
-- Firma:         FRANCO LARA - 11/06/2026 | Agrega Activo y Respuesta para permitir moderacion y replica publica de la atencion del negocio.
CREATE TABLE [dbo].[ReservasUsuariosPublicosResenas](
    [Id] [int] IDENTITY(1,1) NOT NULL,
    [ReservaId] [int] NOT NULL,
    [UsuarioId] [nvarchar](450) NOT NULL,
    [AliasPublico] [nvarchar](120) NOT NULL,
    [Comentario] [nvarchar](800) NOT NULL,
    [Activo] [bit] NOT NULL,
    [Respuesta] [nvarchar](800) NULL,
    [FechaCreacion] [datetime2](7) NOT NULL,
    [UsuarioCreacion] [nvarchar](120) NOT NULL,
CONSTRAINT [PK_ReservasUsuariosPublicosResenas] PRIMARY KEY CLUSTERED ([Id] ASC),
CONSTRAINT [UQ_ReservasUsuariosPublicosResenas_ReservaId] UNIQUE NONCLUSTERED ([ReservaId] ASC)
) ON [PRIMARY];
GO

ALTER TABLE [dbo].[ReservasUsuariosPublicosResenas] ADD CONSTRAINT [DF_ReservasUsuariosPublicosResenas_Activo] DEFAULT ((1)) FOR [Activo];
GO

ALTER TABLE [dbo].[ReservasUsuariosPublicosResenas] ADD CONSTRAINT [DF_ReservasUsuariosPublicosResenas_FechaCreacion] DEFAULT (SYSDATETIME()) FOR [FechaCreacion];
GO

ALTER TABLE [dbo].[ReservasUsuariosPublicosResenas] WITH CHECK ADD CONSTRAINT [FK_ReservasUsuariosPublicosResenas_Reservas_ReservaId]
FOREIGN KEY([ReservaId]) REFERENCES [dbo].[Reservas] ([Id]);
GO
ALTER TABLE [dbo].[ReservasUsuariosPublicosResenas] CHECK CONSTRAINT [FK_ReservasUsuariosPublicosResenas_Reservas_ReservaId];
GO

ALTER TABLE [dbo].[ReservasUsuariosPublicosResenas] WITH CHECK ADD CONSTRAINT [FK_ReservasUsuariosPublicosResenas_AspNetUsers_UsuarioId]
FOREIGN KEY([UsuarioId]) REFERENCES [dbo].[AspNetUsers] ([Id]);
GO
ALTER TABLE [dbo].[ReservasUsuariosPublicosResenas] CHECK CONSTRAINT [FK_ReservasUsuariosPublicosResenas_AspNetUsers_UsuarioId];
GO
