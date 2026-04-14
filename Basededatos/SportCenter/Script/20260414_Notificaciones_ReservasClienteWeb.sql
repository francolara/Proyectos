-- Firma: Codex - 14/04/2026 | CanalOrigen en reservas + campanita de notificaciones para reservas de CLIENTE_WEB.

IF COL_LENGTH('dbo.Reservas', 'CanalOrigen') IS NULL
BEGIN
    ALTER TABLE dbo.Reservas ADD CanalOrigen NVARCHAR(20) NOT NULL CONSTRAINT DF_Reservas_CanalOrigen DEFAULT (N'ADMIN');
END;

IF OBJECT_ID(N'dbo.NegocioNotificaciones', N'U') IS NULL
BEGIN
    CREATE TABLE dbo.NegocioNotificaciones
    (
        Id INT IDENTITY(1,1) NOT NULL CONSTRAINT PK_NegocioNotificaciones PRIMARY KEY,
        NegocioId INT NOT NULL,
        Tipo NVARCHAR(40) NOT NULL,
        Titulo NVARCHAR(120) NOT NULL,
        Mensaje NVARCHAR(300) NOT NULL,
        Entidad NVARCHAR(40) NULL,
        EntidadId INT NULL,
        UrlDestino NVARCHAR(300) NULL,
        Leida BIT NOT NULL CONSTRAINT DF_NegocioNotificaciones_Leida DEFAULT (0),
        FechaRegistroUtc DATETIME2(7) NOT NULL CONSTRAINT DF_NegocioNotificaciones_FechaRegistroUtc DEFAULT (SYSUTCDATETIME()),
        FechaLeidaUtc DATETIME2(7) NULL,
        LeidaPorUserId NVARCHAR(450) NULL,
        CONSTRAINT FK_NegocioNotificaciones_Negocios_NegocioId FOREIGN KEY (NegocioId) REFERENCES dbo.Negocios(Id)
    );
END;

IF NOT EXISTS (
    SELECT 1 FROM sys.indexes
    WHERE object_id = OBJECT_ID(N'dbo.NegocioNotificaciones')
      AND name = N'IX_NegocioNotificaciones_NegocioId_Leida_Fecha'
)
BEGIN
    CREATE INDEX IX_NegocioNotificaciones_NegocioId_Leida_Fecha
        ON dbo.NegocioNotificaciones(NegocioId, Leida, FechaRegistroUtc DESC);
END;