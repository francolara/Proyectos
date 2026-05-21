USE [DbSportCenter]
GO
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO

-- =============================================
-- Author:        FRANCO LARA
-- Create date:   20/05/2026
-- Description:   Tabla de fechas feriadas y tabla de tarifas por feriado (sin DiaSemana) para aplicar precios por rango horario.
-- =============================================

IF OBJECT_ID(N'dbo.Feriados', N'U') IS NULL
BEGIN
    CREATE TABLE dbo.Feriados
    (
        Id INT IDENTITY(1,1) NOT NULL PRIMARY KEY,
        Fecha DATE NOT NULL,
        FechaCreacion DATETIME2 NOT NULL CONSTRAINT DF_Feriados_FechaCreacion DEFAULT (SYSUTCDATETIME()),
        UsuarioCreacion NVARCHAR(200) NULL
    );
END;
GO

IF NOT EXISTS (SELECT 1 FROM sys.indexes WHERE object_id = OBJECT_ID(N'dbo.Feriados') AND name = N'UX_Feriados_Fecha')
BEGIN
    CREATE UNIQUE INDEX UX_Feriados_Fecha ON dbo.Feriados (Fecha);
END;
GO

IF OBJECT_ID(N'dbo.TarifaFeriado', N'U') IS NULL
BEGIN
    CREATE TABLE dbo.TarifaFeriado
    (
        Id INT IDENTITY(1,1) NOT NULL PRIMARY KEY,
        EspacioDeportivoId INT NOT NULL,
        HoraInicio TIME NOT NULL,
        HoraFin TIME NOT NULL,
        Precio DECIMAL(10,2) NOT NULL,
        Activa BIT NOT NULL CONSTRAINT DF_TarifaFeriado_Activa DEFAULT (1),
        FechaCreacion DATETIME2 NOT NULL CONSTRAINT DF_TarifaFeriado_FechaCreacion DEFAULT (SYSUTCDATETIME()),
        UsuarioCreacion NVARCHAR(200) NULL,
        FechaActualizacion DATETIME2 NULL,
        UsuarioActualizacion NVARCHAR(200) NULL,
        CONSTRAINT FK_TarifaFeriado_EspaciosDeportivos_EspacioDeportivoId
            FOREIGN KEY (EspacioDeportivoId) REFERENCES dbo.EspaciosDeportivos(Id)
    );
END;
GO

IF NOT EXISTS (SELECT 1 FROM sys.indexes WHERE object_id = OBJECT_ID(N'dbo.TarifaFeriado') AND name = N'IX_TarifaFeriado_Espacio_Activa_Hora')
BEGIN
    CREATE INDEX IX_TarifaFeriado_Espacio_Activa_Hora
        ON dbo.TarifaFeriado (EspacioDeportivoId, Activa, HoraInicio, HoraFin);
END;
GO
