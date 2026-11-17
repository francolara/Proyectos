-- Firma: FRANCO LARA | Create date: 17/11/2026 | Series de emisión configurables por tipo de comprobante.
USE [SistemaWebHYT];
GO
IF OBJECT_ID(N'dbo.SerieComprobanteHYT',N'U') IS NULL
BEGIN
    CREATE TABLE dbo.SerieComprobanteHYT
    (
        Tipo CHAR(2) NOT NULL CONSTRAINT CK_SerieComprobanteHYT_Tipo CHECK (Tipo IN ('01','03','07')),
        Serie NVARCHAR(30) NOT NULL,
        FechaRegistro DATETIME2(3) NOT NULL DEFAULT SYSUTCDATETIME(),
        UsuarioRegistro NVARCHAR(128) NOT NULL REFERENCES dbo.AspNetUsers(Id),
        CONSTRAINT PK_SerieComprobanteHYT PRIMARY KEY (Tipo,Serie)
    );
END;
GO
