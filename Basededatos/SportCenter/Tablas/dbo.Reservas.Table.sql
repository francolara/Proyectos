USE [DbSportCenter]
GO
/****** Object:  Table [dbo].[Reservas]    Script Date: 3/04/2026 23:17:42 ******/
-- Firma: Codex - 07/04/2026 | Agrega columna Comentario para observaciones de reserva.
-- Firma: Codex - 14/04/2026 | Agrega columna CanalOrigen para identificar reservas de ADMIN o CLIENTE_WEB.
-- Firma: FRANCO LARA - 01/10/2026 | Agrega el correlativo visible de reserva por negocio sin reemplazar el Id tecnico global.
-- Firma: FRANCO LARA - 06/10/2026 | Conserva CodigoMoneda canonico e historico de la operacion.
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE TABLE [dbo].[Reservas](
	[Id] [int] IDENTITY(1,1) NOT NULL,
	[NumeroPorNegocio] [int] NULL,
	[EspacioDeportivoId] [int] NOT NULL,
	[ClienteId] [int] NOT NULL,
	[Fecha] [date] NOT NULL,
	[HoraInicio] [time](7) NOT NULL,
	[HoraFin] [time](7) NOT NULL,
	[Estado] [int] NOT NULL,
	[Total] [decimal](10, 2) NOT NULL,
	[Adelanto] [decimal](10, 2) NOT NULL,
	[Saldo] [decimal](10, 2) NOT NULL,
	[CodigoMoneda] [nvarchar](10) NOT NULL,
	[Comentario] [nvarchar](500) NULL,
	[CanalOrigen] [nvarchar](20) NOT NULL,
	[FechaRegistro] [datetime2](7) NOT NULL,
	[FechaActualizacion] [datetime2](7) NULL,
	[UsuarioActualizacion] [nvarchar](max) NULL,
	[UsuarioCreacion] [nvarchar](max) NULL,
	[RecordatorioEnviado] [bit] NOT NULL,
	[FechaRecordatorio] [datetime2](7) NULL,
 CONSTRAINT [PK_Reservas] PRIMARY KEY CLUSTERED
(
	[Id] ASC
)WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON, OPTIMIZE_FOR_SEQUENTIAL_KEY = OFF) ON [PRIMARY]
) ON [PRIMARY] TEXTIMAGE_ON [PRIMARY]
GO
ALTER TABLE [dbo].[Reservas] ADD  CONSTRAINT [DF_Reservas_CanalOrigen]  DEFAULT (N'ADMIN') FOR [CanalOrigen]
GO
ALTER TABLE [dbo].[Reservas] ADD  CONSTRAINT [DF_Reservas_RecordatorioEnviado]  DEFAULT ((0)) FOR [RecordatorioEnviado]
GO
ALTER TABLE [dbo].[Reservas]  WITH CHECK ADD  CONSTRAINT [FK_Reservas_Clientes_ClienteId] FOREIGN KEY([ClienteId])
REFERENCES [dbo].[Clientes] ([Id])
ON DELETE CASCADE
GO
ALTER TABLE [dbo].[Reservas] CHECK CONSTRAINT [FK_Reservas_Clientes_ClienteId]
GO
ALTER TABLE [dbo].[Reservas]  WITH CHECK ADD  CONSTRAINT [FK_Reservas_EspaciosDeportivos_EspacioDeportivoId] FOREIGN KEY([EspacioDeportivoId])
REFERENCES [dbo].[EspaciosDeportivos] ([Id])
ON DELETE CASCADE
GO
ALTER TABLE [dbo].[Reservas] CHECK CONSTRAINT [FK_Reservas_EspaciosDeportivos_EspacioDeportivoId]
GO
ALTER TABLE [dbo].[Reservas] WITH CHECK ADD CONSTRAINT [FK_Reservas_MonedasSuperMaestro_CodigoMoneda] FOREIGN KEY([CodigoMoneda])
REFERENCES [dbo].[MonedasSuperMaestro] ([Codigo])
GO
ALTER TABLE [dbo].[Reservas] CHECK CONSTRAINT [FK_Reservas_MonedasSuperMaestro_CodigoMoneda]
GO
CREATE NONCLUSTERED INDEX [IX_Reservas_CodigoMoneda] ON [dbo].[Reservas]([CodigoMoneda] ASC)
GO
