-- Firma: FRANCO LARA | Create date: 10/11/2026 | Registro operativo de compras y comprobantes de venta.
USE [SistemaWebHYT];
GO
SET NOCOUNT ON;
IF OBJECT_ID(N'dbo.ComprobanteHYTPago',N'U') IS NULL
BEGIN
    CREATE TABLE dbo.ComprobanteHYTPago
    (
        PagoId BIGINT IDENTITY NOT NULL CONSTRAINT PK_ComprobanteHYTPago PRIMARY KEY,
        ComprobanteId BIGINT NOT NULL REFERENCES dbo.ComprobanteHYT(ComprobanteId),
        EsDetraccion BIT NOT NULL,
        Fecha DATE NOT NULL,
        Importe DECIMAL(18,2) NOT NULL CHECK (Importe > 0),
        Referencia NVARCHAR(100) NOT NULL,
        UsuarioRegistro NVARCHAR(128) NOT NULL REFERENCES dbo.AspNetUsers(Id),
        FechaRegistro DATETIME2(3) NOT NULL DEFAULT SYSUTCDATETIME(),
        CONSTRAINT UQ_ComprobanteHYTPago UNIQUE (ComprobanteId, EsDetraccion, Referencia)
    );
END;
GO
