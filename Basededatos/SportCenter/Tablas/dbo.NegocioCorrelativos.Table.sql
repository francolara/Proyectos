USE [DbSportCenter]
GO
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO

-- Firma: FRANCO LARA - 01/10/2026 | Crea los correlativos visibles por negocio, independientes de los Id tecnicos globales.
CREATE TABLE [dbo].[NegocioCorrelativos](
    [NegocioId] [int] NOT NULL,
    [Entidad] [nvarchar](30) NOT NULL,
    [UltimoNumero] [int] NOT NULL,
    [FechaActualizacion] [datetime2](7) NOT NULL,
    [UsuarioActualizacion] [nvarchar](200) NULL,
    CONSTRAINT [PK_NegocioCorrelativos] PRIMARY KEY CLUSTERED ([NegocioId] ASC, [Entidad] ASC),
    CONSTRAINT [CK_NegocioCorrelativos_Entidad] CHECK ([Entidad] IN (N'RESERVA', N'PAGO', N'CLIENTE', N'SEDE', N'ESPACIO', N'CUPON', N'PROMOCION')),
    CONSTRAINT [CK_NegocioCorrelativos_UltimoNumero] CHECK ([UltimoNumero] >= 0)
) ON [PRIMARY]
GO
ALTER TABLE [dbo].[NegocioCorrelativos] ADD CONSTRAINT [DF_NegocioCorrelativos_FechaActualizacion] DEFAULT (SYSUTCDATETIME()) FOR [FechaActualizacion]
GO
ALTER TABLE [dbo].[NegocioCorrelativos] WITH CHECK ADD CONSTRAINT [FK_NegocioCorrelativos_Negocios_NegocioId] FOREIGN KEY([NegocioId]) REFERENCES [dbo].[Negocios] ([Id]) ON DELETE CASCADE
GO
ALTER TABLE [dbo].[NegocioCorrelativos] CHECK CONSTRAINT [FK_NegocioCorrelativos_Negocios_NegocioId]
GO
