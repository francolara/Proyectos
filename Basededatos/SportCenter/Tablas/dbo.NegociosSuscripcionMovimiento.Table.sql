
GO
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO

-- =============================================
-- Author:        FRANCO LARA
-- Create date:   10/06/2026
-- Firma:         Historial comercial de suscripcion por negocio para registrar activaciones, renovaciones, cambios de plan, extensiones de prueba y ajustes manuales.
-- Firma:         FRANCO LARA - 21/07/2026 | Registra plan comercial y limites anteriores y nuevos de cada aplicacion.
-- =============================================
IF OBJECT_ID(N'dbo.NegociosSuscripcionMovimiento', N'U') IS NULL
BEGIN
    CREATE TABLE [dbo].[NegociosSuscripcionMovimiento]
    (
        [Id] [int] IDENTITY(1,1) NOT NULL,
        [NegocioId] [int] NOT NULL,
        [NegocioSuscripcionId] [int] NULL,
        [TipoMovimiento] [nvarchar](40) NOT NULL,
        [EstadoSuscripcionAnterior] [int] NULL,
        [EstadoSuscripcionNuevo] [int] NOT NULL,
        [EsPruebaAnterior] [bit] NULL,
        [EsPruebaNuevo] [bit] NOT NULL,
        [TipoCobroAnterior] [nvarchar](20) NULL,
        [TipoCobroNuevo] [nvarchar](20) NULL,
        [PlanComercialAnterior] [nvarchar](20) NULL,
        [PlanComercialNuevo] [nvarchar](20) NULL,
        [TipoPlanAnterior] [nvarchar](20) NULL,
        [TipoPlanNuevo] [nvarchar](20) NULL,
        [SedesPermitidasAnterior] [int] NULL,
        [SedesPermitidasNuevo] [int] NULL,
        [EspaciosPermitidosAnterior] [int] NULL,
        [EspaciosPermitidosNuevo] [int] NULL,
        [UsuariosPermitidosAnterior] [int] NULL,
        [UsuariosPermitidosNuevo] [int] NULL,
        [FechaInicioReferencia] [date] NULL,
        [FechaFinReferencia] [date] NULL,
        [DiasGracia] [int] NULL,
        [DiasExtra] [int] NULL,
        [Observacion] [nvarchar](500) NULL,
        [FechaCreacion] [datetime2](7) NOT NULL,
        [UsuarioCreacion] [nvarchar](200) NULL,
        CONSTRAINT [PK_NegociosSuscripcionMovimiento] PRIMARY KEY CLUSTERED ([Id] ASC)
    ) ON [PRIMARY];

    ALTER TABLE [dbo].[NegociosSuscripcionMovimiento]
        ADD CONSTRAINT [DF_NegociosSuscripcionMovimiento_FechaCreacion]
            DEFAULT (SYSUTCDATETIME()) FOR [FechaCreacion];

    ALTER TABLE [dbo].[NegociosSuscripcionMovimiento] WITH CHECK
        ADD CONSTRAINT [FK_NegociosSuscripcionMovimiento_Negocios_NegocioId]
            FOREIGN KEY([NegocioId]) REFERENCES [dbo].[Negocios]([Id]);

    ALTER TABLE [dbo].[NegociosSuscripcionMovimiento] CHECK CONSTRAINT [FK_NegociosSuscripcionMovimiento_Negocios_NegocioId];

    ALTER TABLE [dbo].[NegociosSuscripcionMovimiento] WITH CHECK
        ADD CONSTRAINT [FK_NegociosSuscripcionMovimiento_NegociosSuscripcion_NegocioSuscripcionId]
            FOREIGN KEY([NegocioSuscripcionId]) REFERENCES [dbo].[NegociosSuscripcion]([Id]);

    ALTER TABLE [dbo].[NegociosSuscripcionMovimiento] CHECK CONSTRAINT [FK_NegociosSuscripcionMovimiento_NegociosSuscripcion_NegocioSuscripcionId];

    CREATE NONCLUSTERED INDEX [IX_NegociosSuscripcionMovimiento_Negocio_Fecha]
        ON [dbo].[NegociosSuscripcionMovimiento]([NegocioId] ASC, [FechaCreacion] DESC);
END
GO
