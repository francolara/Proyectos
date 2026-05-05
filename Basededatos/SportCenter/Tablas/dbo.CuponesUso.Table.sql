USE [DbSportCenter]
GO
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE TABLE [dbo].[CuponesUso](
    [Id] [int] IDENTITY(1,1) NOT NULL,
    [CuponId] [int] NOT NULL,
    [ReservaId] [int] NOT NULL,
    [ClienteId] [int] NOT NULL,
    [MontoAntes] [decimal](10,2) NOT NULL,
    [MontoDescuento] [decimal](10,2) NOT NULL,
    [MontoFinal] [decimal](10,2) NOT NULL,
    [CanalOrigen] [nvarchar](20) NOT NULL,
    [FechaUso] [datetime2](7) NOT NULL,
    [UsuarioCreacion] [nvarchar](200) NULL,
CONSTRAINT [PK_CuponesUso] PRIMARY KEY CLUSTERED ([Id] ASC)
) ON [PRIMARY]
GO
ALTER TABLE [dbo].[CuponesUso] ADD CONSTRAINT [DF_CuponesUso_FechaUso] DEFAULT (sysutcdatetime()) FOR [FechaUso]
GO
ALTER TABLE [dbo].[CuponesUso]  WITH CHECK ADD CONSTRAINT [FK_CuponesUso_Cupones_CuponId] FOREIGN KEY([CuponId]) REFERENCES [dbo].[Cupones] ([Id])
GO
ALTER TABLE [dbo].[CuponesUso] CHECK CONSTRAINT [FK_CuponesUso_Cupones_CuponId]
GO
ALTER TABLE [dbo].[CuponesUso]  WITH CHECK ADD CONSTRAINT [FK_CuponesUso_Reservas_ReservaId] FOREIGN KEY([ReservaId]) REFERENCES [dbo].[Reservas] ([Id])
GO
ALTER TABLE [dbo].[CuponesUso] CHECK CONSTRAINT [FK_CuponesUso_Reservas_ReservaId]
GO
CREATE UNIQUE NONCLUSTERED INDEX [UX_CuponesUso_Reserva_Cupon] ON [dbo].[CuponesUso]([ReservaId] ASC, [CuponId] ASC)
GO
