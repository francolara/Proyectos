USE [DbSportCenter]
GO

IF OBJECT_ID(N'dbo.EspacioHorarioAtencion', N'U') IS NULL
BEGIN
    CREATE TABLE [dbo].[EspacioHorarioAtencion](
        [EspacioDeportivoId] [int] NOT NULL,
        [ConfigurarHorarioPorEspacio] [bit] NOT NULL CONSTRAINT [DF_EspacioHorarioAtencion_ConfigurarHorarioPorEspacio] DEFAULT ((0)),
        [AtiendeLunes] [bit] NOT NULL CONSTRAINT [DF_EspacioHorarioAtencion_Lunes] DEFAULT ((1)),
        [AtiendeMartes] [bit] NOT NULL CONSTRAINT [DF_EspacioHorarioAtencion_Martes] DEFAULT ((1)),
        [AtiendeMiercoles] [bit] NOT NULL CONSTRAINT [DF_EspacioHorarioAtencion_Miercoles] DEFAULT ((1)),
        [AtiendeJueves] [bit] NOT NULL CONSTRAINT [DF_EspacioHorarioAtencion_Jueves] DEFAULT ((1)),
        [AtiendeViernes] [bit] NOT NULL CONSTRAINT [DF_EspacioHorarioAtencion_Viernes] DEFAULT ((1)),
        [AtiendeSabado] [bit] NOT NULL CONSTRAINT [DF_EspacioHorarioAtencion_Sabado] DEFAULT ((1)),
        [AtiendeDomingo] [bit] NOT NULL CONSTRAINT [DF_EspacioHorarioAtencion_Domingo] DEFAULT ((1)),
        [HoraApertura] [time](7) NOT NULL CONSTRAINT [DF_EspacioHorarioAtencion_HoraApertura] DEFAULT ('08:00'),
        [HoraCierre] [time](7) NOT NULL CONSTRAINT [DF_EspacioHorarioAtencion_HoraCierre] DEFAULT ('23:00'),
        [FechaCreacion] [datetime2](7) NOT NULL CONSTRAINT [DF_EspacioHorarioAtencion_FechaCreacion] DEFAULT (SYSUTCDATETIME()),
        [UsuarioCreacion] [nvarchar](200) NULL,
        [FechaActualizacion] [datetime2](7) NULL,
        [UsuarioActualizacion] [nvarchar](200) NULL,
        CONSTRAINT [PK_EspacioHorarioAtencion] PRIMARY KEY CLUSTERED ([EspacioDeportivoId] ASC),
        CONSTRAINT [FK_EspacioHorarioAtencion_Espacio] FOREIGN KEY([EspacioDeportivoId]) REFERENCES [dbo].[EspaciosDeportivos]([Id])
    );
END
GO

-- =============================================
-- Author:        FRANCO LARA
-- Create date:   26/05/2026
-- Description:   Crea tabla de horario configurable por espacio deportivo.
-- =============================================
