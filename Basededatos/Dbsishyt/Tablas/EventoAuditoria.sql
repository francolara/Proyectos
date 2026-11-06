-- =============================================
-- Author:        FRANCO LARA
-- Create date:   03/11/2026
-- Description:   Define EventoAuditoria para la base inicial de HYT.
-- =============================================
USE [SistemaWebHYT];
GO
SET NOCOUNT ON;
IF OBJECT_ID(N'dbo.EventoAuditoria', N'U') IS NULL
BEGIN
    CREATE TABLE dbo.EventoAuditoria
    (
        EventoId BIGINT IDENTITY NOT NULL CONSTRAINT PK_EventoAuditoria PRIMARY KEY,
        CorrelacionId UNIQUEIDENTIFIER NOT NULL,
        Entidad NVARCHAR(80) NOT NULL,
        RegistroId NVARCHAR(100) NOT NULL,
        Accion NVARCHAR(60) NOT NULL,
        EstadoAnterior NVARCHAR(40) NULL,
        EstadoNuevo NVARCHAR(40) NULL,
        UsuarioId NVARCHAR(128) NOT NULL,
        FechaRegistro DATETIME2(3) NOT NULL DEFAULT SYSUTCDATETIME(),
        FechaEvento DATETIME2(3) NULL,
        Motivo NVARCHAR(1000) NULL,
        CambiosJson NVARCHAR(MAX) NULL,
        CONSTRAINT FK_EventoAuditoria_Usuario FOREIGN KEY (UsuarioId) REFERENCES dbo.AspNetUsers(Id),
        CONSTRAINT CK_EventoAuditoria_Json CHECK (CambiosJson IS NULL OR ISJSON(CambiosJson) = 1)
    );
END;
GO
IF NOT EXISTS (SELECT 1 FROM sys.indexes WHERE object_id = OBJECT_ID(N'dbo.EventoAuditoria') AND name = N'IX_EventoAuditoria_Registro')
    CREATE INDEX IX_EventoAuditoria_Registro ON dbo.EventoAuditoria(Entidad, RegistroId, FechaRegistro, EventoId);
GO
