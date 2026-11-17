-- Firma: FRANCO LARA | Create date: 17/11/2026 | Conserva la referencia emitida retirada de la tabla operativa.
USE [SistemaWebHYT];
GO
SET NOCOUNT ON;
IF OBJECT_ID(N'dbo.ReferenciaComprobanteHistorial',N'U') IS NULL
BEGIN
    CREATE TABLE dbo.ReferenciaComprobanteHistorial
    (
        ReferenciaId BIGINT NOT NULL CONSTRAINT PK_ReferenciaComprobanteHistorial PRIMARY KEY,
        OrdenCompraId BIGINT NOT NULL REFERENCES dbo.OrdenCompraCliente(OrdenCompraId),
        ComprobanteId BIGINT NOT NULL,
        FechaDocumento DATE NOT NULL,
        CodigoTipoDocumento NVARCHAR(2) NOT NULL,
        Serie NVARCHAR(30) NOT NULL,
        NumeroDocumento NVARCHAR(50) NOT NULL,
        FechaRegistro DATETIME2(3) NOT NULL,
        UsuarioRegistro NVARCHAR(128) NOT NULL REFERENCES dbo.AspNetUsers(Id),
        FechaRetiro DATETIME2(3) NOT NULL,
        MotivoRetiro NVARCHAR(1000) NOT NULL
    );
    CREATE INDEX IX_ReferenciaComprobanteHistorial_Orden
        ON dbo.ReferenciaComprobanteHistorial(OrdenCompraId,ReferenciaId DESC);
END;
GO
