USE [DbSportCenter]
GO
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
-- =============================================
-- Author:        FRANCO LARA
-- Create date:   26/05/2026
-- Description:   Estado de avance del onboarding inicial por negocio.
-- =============================================
-- Firma: Codex - 26/05/2026 | Se crea tabla para persistir paso actual y estado de finalizacion del wizard de onboarding por negocio.
CREATE TABLE [dbo].[NegocioOnboardingEstado](
    [NegocioId] [int] NOT NULL,
    [PasoActual] [tinyint] NOT NULL,
    [Completado] [bit] NOT NULL,
    [FechaUltimaActualizacionUtc] [datetime2](7) NOT NULL,
    [UsuarioUltimaActualizacion] [nvarchar](200) NULL,
    [FechaCompletadoUtc] [datetime2](7) NULL,
    [UsuarioCompletado] [nvarchar](200) NULL,
 CONSTRAINT [PK_NegocioOnboardingEstado] PRIMARY KEY CLUSTERED
(
    [NegocioId] ASC
)WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON) ON [PRIMARY]
) ON [PRIMARY]
GO
ALTER TABLE [dbo].[NegocioOnboardingEstado] ADD CONSTRAINT [DF_NegocioOnboardingEstado_PasoActual] DEFAULT ((1)) FOR [PasoActual]
GO
ALTER TABLE [dbo].[NegocioOnboardingEstado] ADD CONSTRAINT [DF_NegocioOnboardingEstado_Completado] DEFAULT ((0)) FOR [Completado]
GO
ALTER TABLE [dbo].[NegocioOnboardingEstado] ADD CONSTRAINT [DF_NegocioOnboardingEstado_FechaUltimaActualizacionUtc] DEFAULT (SYSUTCDATETIME()) FOR [FechaUltimaActualizacionUtc]
GO
ALTER TABLE [dbo].[NegocioOnboardingEstado]  WITH CHECK ADD  CONSTRAINT [FK_NegocioOnboardingEstado_Negocios_NegocioId] FOREIGN KEY([NegocioId])
REFERENCES [dbo].[Negocios] ([Id])
ON DELETE CASCADE
GO
ALTER TABLE [dbo].[NegocioOnboardingEstado] CHECK CONSTRAINT [FK_NegocioOnboardingEstado_Negocios_NegocioId]
GO
