
GO
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO

-- =============================================
-- Author:        FRANCO LARA
-- Create date:   10/06/2026
-- Firma:         Registro manual y futuro conciliable de cobros de suscripcion por negocio.
-- Firma:         FRANCO LARA - 21/07/2026 | Guarda plan comercial y limites objetivo aplicados con el cobro.
-- =============================================
IF OBJECT_ID(N'dbo.NegociosSuscripcionPago', N'U') IS NULL
BEGIN
    CREATE TABLE [dbo].[NegociosSuscripcionPago]
    (
        [Id] [int] IDENTITY(1,1) NOT NULL,
        [NegocioId] [int] NOT NULL,
        [NegocioSuscripcionId] [int] NOT NULL,
        [NegocioSuscripcionMovimientoId] [int] NULL,
        [TipoPago] [nvarchar](30) NOT NULL,
        [EstadoPago] [nvarchar](20) NOT NULL,
        [Monto] [decimal](12, 2) NOT NULL,
        [Moneda] [nvarchar](10) NOT NULL,
        [FechaPago] [datetime2](7) NOT NULL,
        [FechaVencimiento] [date] NULL,
        [OperacionNumero] [nvarchar](100) NULL,
        [EntidadFinanciera] [nvarchar](120) NULL,
        [ReferenciaExterna] [nvarchar](120) NULL,
        [AccionAplicacion] [nvarchar](30) NULL,
        [AplicarAlConfirmar] [bit] NOT NULL,
        [AplicadoSuscripcion] [bit] NOT NULL,
        [FechaAplicacion] [datetime2](7) NULL,
        [UsuarioAplicacion] [nvarchar](200) NULL,
        [TipoCobroObjetivo] [nvarchar](20) NULL,
        [PlanComercialObjetivo] [nvarchar](20) NULL,
        [TipoPlanObjetivo] [nvarchar](20) NULL,
        [SedesPermitidasObjetivo] [int] NULL,
        [EspaciosPermitidosObjetivo] [int] NULL,
        [UsuariosPermitidosObjetivo] [int] NULL,
        [FechaInicioPlanObjetivo] [date] NULL,
        [DiasGraciaObjetivo] [int] NULL,
        [Observacion] [nvarchar](500) NULL,
        [FechaCreacion] [datetime2](7) NOT NULL,
        [UsuarioCreacion] [nvarchar](200) NULL,
        [FechaActualizacion] [datetime2](7) NULL,
        [UsuarioActualizacion] [nvarchar](200) NULL,
        CONSTRAINT [PK_NegociosSuscripcionPago] PRIMARY KEY CLUSTERED ([Id] ASC)
    ) ON [PRIMARY];

    ALTER TABLE [dbo].[NegociosSuscripcionPago]
        ADD CONSTRAINT [DF_NegociosSuscripcionPago_FechaCreacion]
            DEFAULT (SYSUTCDATETIME()) FOR [FechaCreacion];

    ALTER TABLE [dbo].[NegociosSuscripcionPago]
        ADD CONSTRAINT [DF_NegociosSuscripcionPago_Moneda]
            DEFAULT (N'PEN') FOR [Moneda];

    ALTER TABLE [dbo].[NegociosSuscripcionPago]
        ADD CONSTRAINT [DF_NegociosSuscripcionPago_AplicarAlConfirmar]
            DEFAULT ((0)) FOR [AplicarAlConfirmar];

    ALTER TABLE [dbo].[NegociosSuscripcionPago]
        ADD CONSTRAINT [DF_NegociosSuscripcionPago_AplicadoSuscripcion]
            DEFAULT ((0)) FOR [AplicadoSuscripcion];

    ALTER TABLE [dbo].[NegociosSuscripcionPago] WITH CHECK
        ADD CONSTRAINT [FK_NegociosSuscripcionPago_Negocios_NegocioId]
            FOREIGN KEY([NegocioId]) REFERENCES [dbo].[Negocios]([Id]);

    ALTER TABLE [dbo].[NegociosSuscripcionPago] CHECK CONSTRAINT [FK_NegociosSuscripcionPago_Negocios_NegocioId];

    ALTER TABLE [dbo].[NegociosSuscripcionPago] WITH CHECK
        ADD CONSTRAINT [FK_NegociosSuscripcionPago_NegociosSuscripcion_NegocioSuscripcionId]
            FOREIGN KEY([NegocioSuscripcionId]) REFERENCES [dbo].[NegociosSuscripcion]([Id]);

    ALTER TABLE [dbo].[NegociosSuscripcionPago] CHECK CONSTRAINT [FK_NegociosSuscripcionPago_NegociosSuscripcion_NegocioSuscripcionId];

    ALTER TABLE [dbo].[NegociosSuscripcionPago] WITH CHECK
        ADD CONSTRAINT [FK_NegociosSuscripcionPago_NegociosSuscripcionMovimiento_NegocioSuscripcionMovimientoId]
            FOREIGN KEY([NegocioSuscripcionMovimientoId]) REFERENCES [dbo].[NegociosSuscripcionMovimiento]([Id]);

    ALTER TABLE [dbo].[NegociosSuscripcionPago] CHECK CONSTRAINT [FK_NegociosSuscripcionPago_NegociosSuscripcionMovimiento_NegocioSuscripcionMovimientoId];

    CREATE NONCLUSTERED INDEX [IX_NegociosSuscripcionPago_Negocio_Fecha]
        ON [dbo].[NegociosSuscripcionPago]([NegocioId] ASC, [FechaPago] DESC, [Id] DESC);
END
GO
