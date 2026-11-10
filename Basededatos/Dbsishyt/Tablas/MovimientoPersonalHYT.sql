-- =============================================
-- Author:        FRANCO LARA
-- Create date:   09/11/2026
-- Description:   Define dbo.MovimientoPersonalHYT para las etapas 3 y 4 de Planillas.
-- =============================================
USE [SistemaWebHYT];
GO
IF OBJECT_ID(N'dbo.MovimientoPersonalHYT', N'U') IS NULL
BEGIN
    CREATE TABLE dbo.MovimientoPersonalHYT
    (
        MovimientoId BIGINT IDENTITY NOT NULL CONSTRAINT PK_MovimientoPersonalHYT PRIMARY KEY,
        TrabajadorId BIGINT NOT NULL CONSTRAINT FK_MovimientoPersonalHYT_Trabajador REFERENCES dbo.TrabajadorHYT(TrabajadorId),
        CeseId BIGINT NULL CONSTRAINT FK_MovimientoPersonalHYT_Cese REFERENCES dbo.CeseTrabajadorHYT(CeseId),
        TipoTrabajador VARCHAR(10) NOT NULL CONSTRAINT CK_MovimientoPersonalHYT_Trabajador CHECK (TipoTrabajador IN ('PLANILLA', 'EXTERNO')),
        Concepto VARCHAR(24) NOT NULL CONSTRAINT CK_MovimientoPersonalHYT_Concepto CHECK
            (Concepto IN ('REMUNERACION', 'SERVICIO', 'GRATIFICACION', 'LIQUIDACION', 'VACACIONES_PENDIENTES',
                          'ADELANTO', 'ADELANTO_GRATIFICACION', 'DESCUENTO', 'PAGO')),
        Fecha DATE NOT NULL,
        Descripcion NVARCHAR(300) NOT NULL,
        Importe DECIMAL(18,2) NOT NULL CONSTRAINT CK_MovimientoPersonalHYT_Importe CHECK (Importe > 0),
        MonedaCodigo NVARCHAR(3) NOT NULL CONSTRAINT FK_MovimientoPersonalHYT_Moneda REFERENCES dbo.Moneda(Codigo),
        Frecuencia VARCHAR(10) NULL,
        PeriodoDesde DATE NULL,
        PeriodoHasta DATE NULL,
        Referencia NVARCHAR(100) NULL,
        OperacionId UNIQUEIDENTIFIER NOT NULL CONSTRAINT UQ_MovimientoPersonalHYT_Operacion UNIQUE,
        FechaRegistro DATETIME2(3) NOT NULL CONSTRAINT DF_MovimientoPersonalHYT_Fecha DEFAULT SYSUTCDATETIME(),
        UsuarioId NVARCHAR(128) NOT NULL CONSTRAINT FK_MovimientoPersonalHYT_Usuario REFERENCES dbo.AspNetUsers(Id),
        CONSTRAINT CK_MovimientoPersonalHYT_Periodo CHECK (
            (Frecuencia IS NULL AND PeriodoDesde IS NULL AND PeriodoHasta IS NULL)
            OR (Frecuencia IS NOT NULL AND Frecuencia IN ('SEMANAL', 'QUINCENAL', 'MENSUAL')
                AND PeriodoDesde IS NOT NULL AND PeriodoHasta IS NOT NULL AND PeriodoDesde <= PeriodoHasta))
    );
END;
GO
