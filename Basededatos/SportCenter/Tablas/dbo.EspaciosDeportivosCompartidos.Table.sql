GO
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE TABLE [dbo].[EspaciosDeportivosCompartidos](
    [Id] [int] IDENTITY(1,1) NOT NULL,
    [EspacioDeportivoId] [int] NOT NULL,
    [EspacioRelacionadoId] [int] NOT NULL,
    [TipoRelacion] [nvarchar](40) NOT NULL,
    [Activo] [bit] NOT NULL,
    [FechaCreacion] [datetime2](7) NOT NULL,
    [UsuarioCreacion] [nvarchar](200) NULL,
    [FechaActualizacion] [datetime2](7) NULL,
    [UsuarioActualizacion] [nvarchar](200) NULL,
PRIMARY KEY CLUSTERED
(
    [Id] ASC
)
) ON [PRIMARY]
GO
ALTER TABLE [dbo].[EspaciosDeportivosCompartidos] ADD CONSTRAINT [DF_EspaciosDeportivosCompartidos_TipoRelacion] DEFAULT (N'DIRECTO') FOR [TipoRelacion]
GO
ALTER TABLE [dbo].[EspaciosDeportivosCompartidos] ADD CONSTRAINT [DF_EspaciosDeportivosCompartidos_Activo] DEFAULT ((1)) FOR [Activo]
GO
ALTER TABLE [dbo].[EspaciosDeportivosCompartidos] ADD CONSTRAINT [DF_EspaciosDeportivosCompartidos_FechaCreacion] DEFAULT (sysutcdatetime()) FOR [FechaCreacion]
GO
ALTER TABLE [dbo].[EspaciosDeportivosCompartidos]  WITH CHECK ADD CONSTRAINT [CK_EspaciosDeportivosCompartidos_TipoRelacion] CHECK  (([TipoRelacion]=N'DIRECTO' OR [TipoRelacion]=N'COMPUESTO_COMPONENTE'))
GO
ALTER TABLE [dbo].[EspaciosDeportivosCompartidos] CHECK CONSTRAINT [CK_EspaciosDeportivosCompartidos_TipoRelacion]
GO
ALTER TABLE [dbo].[EspaciosDeportivosCompartidos]  WITH CHECK ADD CONSTRAINT [FK_EspaciosDeportivosCompartidos_Espacio] FOREIGN KEY([EspacioDeportivoId])
REFERENCES [dbo].[EspaciosDeportivos] ([Id])
GO
ALTER TABLE [dbo].[EspaciosDeportivosCompartidos] CHECK CONSTRAINT [FK_EspaciosDeportivosCompartidos_Espacio]
GO
ALTER TABLE [dbo].[EspaciosDeportivosCompartidos]  WITH CHECK ADD CONSTRAINT [FK_EspaciosDeportivosCompartidos_EspacioRelacionado] FOREIGN KEY([EspacioRelacionadoId])
REFERENCES [dbo].[EspaciosDeportivos] ([Id])
GO
ALTER TABLE [dbo].[EspaciosDeportivosCompartidos] CHECK CONSTRAINT [FK_EspaciosDeportivosCompartidos_EspacioRelacionado]
GO
CREATE UNIQUE NONCLUSTERED INDEX [UX_EspaciosDeportivosCompartidos_ParActivo]
ON [dbo].[EspaciosDeportivosCompartidos] ([EspacioDeportivoId] ASC, [EspacioRelacionadoId] ASC)
WHERE [Activo] = (1)
GO

-- =============================================
-- Author:        FRANCO LARA
-- Create date:   06/06/2026
-- Description:   Relacion bidireccional operativa entre espacios deportivos para bloqueo compartido de horarios.
-- =============================================
-- =============================================
-- Author:        FRANCO LARA
-- Create date:   08/06/2026
-- Description:   Evoluciona la relacion operativa para distinguir bloqueo directo y espacios compuestos por componentes.
-- =============================================
