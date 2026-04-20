USE [DbSportCenter]
GO
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
-- =============================================
-- Author:        FRANCO LARA
-- Create date:   18/04/2026
-- Description:   Tabla transaccional para desafios entre usuarios publicos.
-- =============================================
CREATE TABLE [dbo].[Desafio](
    [Id] [int] IDENTITY(1,1) NOT NULL,
    [IdUsuarioRetador] [nvarchar](450) NOT NULL,
    [IdUsuarioRetado] [nvarchar](450) NOT NULL,
    [IdDeporte] [int] NOT NULL,
    [IdNivel] [int] NOT NULL,
    [Distrito] [char](6) NOT NULL,
    [FechaTentativa] [date] NOT NULL,
    [HoraTentativa] [time](7) NOT NULL,
    [CanchaSugerida] [nvarchar](150) NULL,
    [Modalidad] [nvarchar](120) NOT NULL,
    [Mensaje] [nvarchar](500) NULL,
    [FormaPago] [nvarchar](120) NOT NULL,
    [Estado] [nvarchar](20) NOT NULL,
    [FechaCreacion] [datetime2](7) NOT NULL,
    [FechaRespuesta] [datetime2](7) NULL,
    [Activo] [bit] NOT NULL,
    [UsuarioCreacion] [nvarchar](120) NOT NULL,
    [UsuarioActualizacion] [nvarchar](120) NULL,
CONSTRAINT [PK_Desafio] PRIMARY KEY CLUSTERED ([Id] ASC)
) ON [PRIMARY];
GO
ALTER TABLE [dbo].[Desafio] ADD CONSTRAINT [DF_Desafio_Estado] DEFAULT (N'Pendiente') FOR [Estado];
GO
ALTER TABLE [dbo].[Desafio] ADD CONSTRAINT [DF_Desafio_Activo] DEFAULT ((1)) FOR [Activo];
GO
ALTER TABLE [dbo].[Desafio] ADD CONSTRAINT [DF_Desafio_FechaCreacion] DEFAULT (SYSDATETIME()) FOR [FechaCreacion];
GO
ALTER TABLE [dbo].[Desafio]  WITH CHECK ADD CONSTRAINT [CK_Desafio_Estado]
CHECK ([Estado] IN (N'Pendiente', N'Aceptado', N'Rechazado', N'Cancelado', N'Finalizado'));
GO
ALTER TABLE [dbo].[Desafio]  WITH CHECK ADD CONSTRAINT [CK_Desafio_Usuarios_Diferentes]
CHECK ([IdUsuarioRetador] <> [IdUsuarioRetado]);
GO
ALTER TABLE [dbo].[Desafio]  WITH CHECK ADD CONSTRAINT [FK_Desafio_AspNetUsers_IdUsuarioRetador]
FOREIGN KEY([IdUsuarioRetador]) REFERENCES [dbo].[AspNetUsers] ([Id]);
GO
ALTER TABLE [dbo].[Desafio] CHECK CONSTRAINT [FK_Desafio_AspNetUsers_IdUsuarioRetador];
GO
ALTER TABLE [dbo].[Desafio]  WITH CHECK ADD CONSTRAINT [FK_Desafio_AspNetUsers_IdUsuarioRetado]
FOREIGN KEY([IdUsuarioRetado]) REFERENCES [dbo].[AspNetUsers] ([Id]);
GO
ALTER TABLE [dbo].[Desafio] CHECK CONSTRAINT [FK_Desafio_AspNetUsers_IdUsuarioRetado];
GO
ALTER TABLE [dbo].[Desafio]  WITH CHECK ADD CONSTRAINT [FK_Desafio_TiposDeporte_IdDeporte]
FOREIGN KEY([IdDeporte]) REFERENCES [dbo].[TiposDeporte] ([Id]);
GO
ALTER TABLE [dbo].[Desafio] CHECK CONSTRAINT [FK_Desafio_TiposDeporte_IdDeporte];
GO
ALTER TABLE [dbo].[Desafio]  WITH CHECK ADD CONSTRAINT [FK_Desafio_NivelDesafio_IdNivel]
FOREIGN KEY([IdNivel]) REFERENCES [dbo].[NivelDesafio] ([IdNivel]);
GO
ALTER TABLE [dbo].[Desafio] CHECK CONSTRAINT [FK_Desafio_NivelDesafio_IdNivel];
GO
ALTER TABLE [dbo].[Desafio]  WITH CHECK ADD CONSTRAINT [FK_Desafio_UbigeoDistritos_Distrito]
FOREIGN KEY([Distrito]) REFERENCES [dbo].[UbigeoDistritos] ([CodigoUbigeo]);
GO
ALTER TABLE [dbo].[Desafio] CHECK CONSTRAINT [FK_Desafio_UbigeoDistritos_Distrito];
GO
CREATE NONCLUSTERED INDEX [IX_Desafio_Retador_Estado]
    ON [dbo].[Desafio]([IdUsuarioRetador] ASC, [Estado] ASC, [FechaCreacion] DESC);
GO
CREATE NONCLUSTERED INDEX [IX_Desafio_Retado_Estado]
    ON [dbo].[Desafio]([IdUsuarioRetado] ASC, [Estado] ASC, [FechaCreacion] DESC);
GO
