
GO
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE TABLE [dbo].[EspacioHorarioAtencion](
    [EspacioDeportivoId] [int] NOT NULL,
    [ConfigurarHorarioPorEspacio] [bit] NOT NULL,
    [AtiendeLunes] [bit] NOT NULL,
    [AtiendeMartes] [bit] NOT NULL,
    [AtiendeMiercoles] [bit] NOT NULL,
    [AtiendeJueves] [bit] NOT NULL,
    [AtiendeViernes] [bit] NOT NULL,
    [AtiendeSabado] [bit] NOT NULL,
    [AtiendeDomingo] [bit] NOT NULL,
    [HoraApertura] [time](7) NOT NULL,
    [HoraCierre] [time](7) NOT NULL,
    [FechaCreacion] [datetime2](7) NOT NULL,
    [UsuarioCreacion] [nvarchar](200) NULL,
    [FechaActualizacion] [datetime2](7) NULL,
    [UsuarioActualizacion] [nvarchar](200) NULL,
PRIMARY KEY CLUSTERED
(
    [EspacioDeportivoId] ASC
)
) ON [PRIMARY]
GO
ALTER TABLE [dbo].[EspacioHorarioAtencion] ADD CONSTRAINT [DF_EspacioHorarioAtencion_ConfigurarHorarioPorEspacio] DEFAULT ((0)) FOR [ConfigurarHorarioPorEspacio]
GO
ALTER TABLE [dbo].[EspacioHorarioAtencion] ADD CONSTRAINT [DF_EspacioHorarioAtencion_Lunes] DEFAULT ((1)) FOR [AtiendeLunes]
GO
ALTER TABLE [dbo].[EspacioHorarioAtencion] ADD CONSTRAINT [DF_EspacioHorarioAtencion_Martes] DEFAULT ((1)) FOR [AtiendeMartes]
GO
ALTER TABLE [dbo].[EspacioHorarioAtencion] ADD CONSTRAINT [DF_EspacioHorarioAtencion_Miercoles] DEFAULT ((1)) FOR [AtiendeMiercoles]
GO
ALTER TABLE [dbo].[EspacioHorarioAtencion] ADD CONSTRAINT [DF_EspacioHorarioAtencion_Jueves] DEFAULT ((1)) FOR [AtiendeJueves]
GO
ALTER TABLE [dbo].[EspacioHorarioAtencion] ADD CONSTRAINT [DF_EspacioHorarioAtencion_Viernes] DEFAULT ((1)) FOR [AtiendeViernes]
GO
ALTER TABLE [dbo].[EspacioHorarioAtencion] ADD CONSTRAINT [DF_EspacioHorarioAtencion_Sabado] DEFAULT ((1)) FOR [AtiendeSabado]
GO
ALTER TABLE [dbo].[EspacioHorarioAtencion] ADD CONSTRAINT [DF_EspacioHorarioAtencion_Domingo] DEFAULT ((1)) FOR [AtiendeDomingo]
GO
ALTER TABLE [dbo].[EspacioHorarioAtencion] ADD CONSTRAINT [DF_EspacioHorarioAtencion_HoraApertura] DEFAULT ('08:00') FOR [HoraApertura]
GO
ALTER TABLE [dbo].[EspacioHorarioAtencion] ADD CONSTRAINT [DF_EspacioHorarioAtencion_HoraCierre] DEFAULT ('23:00') FOR [HoraCierre]
GO
ALTER TABLE [dbo].[EspacioHorarioAtencion] ADD CONSTRAINT [DF_EspacioHorarioAtencion_FechaCreacion] DEFAULT (sysutcdatetime()) FOR [FechaCreacion]
GO
ALTER TABLE [dbo].[EspacioHorarioAtencion]  WITH CHECK ADD CONSTRAINT [FK_EspacioHorarioAtencion_Espacio] FOREIGN KEY([EspacioDeportivoId])
REFERENCES [dbo].[EspaciosDeportivos] ([Id])
GO
ALTER TABLE [dbo].[EspacioHorarioAtencion] CHECK CONSTRAINT [FK_EspacioHorarioAtencion_Espacio]
GO

-- =============================================
-- Author:        FRANCO LARA
-- Create date:   26/05/2026
-- Description:   Horario de atencion configurable por espacio deportivo.
-- =============================================
