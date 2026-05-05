USE [DbSportCenter]
GO
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE TABLE [dbo].[Cupones](
    [Id] [int] IDENTITY(1,1) NOT NULL,
    [NegocioId] [int] NOT NULL,
    [SedeId] [int] NULL,
    [EspacioDeportivoId] [int] NULL,
    [CodigoCupon] [nvarchar](30) NOT NULL,
    [Nombre] [nvarchar](150) NOT NULL,
    [TipoDescuento] [nvarchar](20) NOT NULL,
    [ValorDescuento] [decimal](10,2) NOT NULL,
    [CantidadMaxUsos] [int] NOT NULL,
    [CantidadUsosActuales] [int] NOT NULL,
    [FechaInicio] [date] NOT NULL,
    [FechaFin] [date] NOT NULL,
    [Activo] [bit] NOT NULL,
    [FechaRegistro] [datetime2](7) NOT NULL,
    [UsuarioCreacion] [nvarchar](200) NULL,
    [FechaActualizacion] [datetime2](7) NULL,
    [UsuarioActualizacion] [nvarchar](200) NULL,
CONSTRAINT [PK_Cupones] PRIMARY KEY CLUSTERED ([Id] ASC)
) ON [PRIMARY]
GO
ALTER TABLE [dbo].[Cupones] ADD CONSTRAINT [DF_Cupones_Activo] DEFAULT ((1)) FOR [Activo]
GO
ALTER TABLE [dbo].[Cupones] ADD CONSTRAINT [DF_Cupones_CantidadUsosActuales] DEFAULT ((0)) FOR [CantidadUsosActuales]
GO
ALTER TABLE [dbo].[Cupones] ADD CONSTRAINT [DF_Cupones_FechaRegistro] DEFAULT (sysutcdatetime()) FOR [FechaRegistro]
GO
ALTER TABLE [dbo].[Cupones]  WITH CHECK ADD CONSTRAINT [FK_Cupones_Negocios_NegocioId] FOREIGN KEY([NegocioId]) REFERENCES [dbo].[Negocios] ([Id])
GO
ALTER TABLE [dbo].[Cupones] CHECK CONSTRAINT [FK_Cupones_Negocios_NegocioId]
GO
ALTER TABLE [dbo].[Cupones]  WITH CHECK ADD CONSTRAINT [FK_Cupones_Sedes_SedeId] FOREIGN KEY([SedeId]) REFERENCES [dbo].[Sedes] ([Id])
GO
ALTER TABLE [dbo].[Cupones] CHECK CONSTRAINT [FK_Cupones_Sedes_SedeId]
GO
ALTER TABLE [dbo].[Cupones]  WITH CHECK ADD CONSTRAINT [FK_Cupones_Espacios_EspacioDeportivoId] FOREIGN KEY([EspacioDeportivoId]) REFERENCES [dbo].[EspaciosDeportivos] ([Id])
GO
ALTER TABLE [dbo].[Cupones] CHECK CONSTRAINT [FK_Cupones_Espacios_EspacioDeportivoId]
GO
CREATE UNIQUE NONCLUSTERED INDEX [UX_Cupones_Negocio_Codigo] ON [dbo].[Cupones]([NegocioId] ASC, [CodigoCupon] ASC)
GO
