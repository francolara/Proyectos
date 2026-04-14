USE [DbSportCenter]
GO
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO

-- Firma: Codex - 14/04/2026 | Tabla de notificaciones por negocio para campanita admin de reservas web.
CREATE TABLE [dbo].[NegocioNotificaciones](
    [Id] [int] IDENTITY(1,1) NOT NULL,
    [NegocioId] [int] NOT NULL,
    [Tipo] [nvarchar](40) NOT NULL,
    [Titulo] [nvarchar](120) NOT NULL,
    [Mensaje] [nvarchar](300) NOT NULL,
    [Entidad] [nvarchar](40) NULL,
    [EntidadId] [int] NULL,
    [UrlDestino] [nvarchar](300) NULL,
    [Leida] [bit] NOT NULL,
    [FechaRegistroUtc] [datetime2](7) NOT NULL,
    [FechaLeidaUtc] [datetime2](7) NULL,
    [LeidaPorUserId] [nvarchar](450) NULL,
 CONSTRAINT [PK_NegocioNotificaciones] PRIMARY KEY CLUSTERED
(
    [Id] ASC
)
) ON [PRIMARY]
GO

ALTER TABLE [dbo].[NegocioNotificaciones]
ADD CONSTRAINT [DF_NegocioNotificaciones_Leida] DEFAULT ((0)) FOR [Leida]
GO

ALTER TABLE [dbo].[NegocioNotificaciones]
ADD CONSTRAINT [DF_NegocioNotificaciones_FechaRegistroUtc] DEFAULT (SYSUTCDATETIME()) FOR [FechaRegistroUtc]
GO

ALTER TABLE [dbo].[NegocioNotificaciones] WITH CHECK
ADD CONSTRAINT [FK_NegocioNotificaciones_Negocios_NegocioId] FOREIGN KEY([NegocioId])
REFERENCES [dbo].[Negocios] ([Id])
GO

ALTER TABLE [dbo].[NegocioNotificaciones] CHECK CONSTRAINT [FK_NegocioNotificaciones_Negocios_NegocioId]
GO

CREATE INDEX [IX_NegocioNotificaciones_NegocioId_Leida_Fecha]
ON [dbo].[NegocioNotificaciones] ([NegocioId], [Leida], [FechaRegistroUtc] DESC)
GO
