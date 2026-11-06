-- =============================================
-- Author:        FRANCO LARA
-- Create date:   03/11/2026
-- Description:   Define OrdenCompraCliente para la base inicial de HYT.
-- =============================================
USE [SistemaWebHYT];
GO
SET NOCOUNT ON;
IF OBJECT_ID(N'dbo.OrdenCompraCliente', N'U') IS NULL
BEGIN
    CREATE TABLE dbo.OrdenCompraCliente
    (
        OrdenCompraId BIGINT IDENTITY NOT NULL CONSTRAINT PK_OrdenCompraCliente PRIMARY KEY,
        NumeroRegistroInterno NVARCHAR(40) NOT NULL CONSTRAINT UQ_OC_Interno UNIQUE,
        NumeroOCCliente NVARCHAR(100) NOT NULL,
        PresupuestoId BIGINT NOT NULL CONSTRAINT UQ_OC_Presupuesto UNIQUE,
        RevisionId BIGINT NOT NULL,
        FechaEmisionOC DATE NOT NULL,
        FechaRecepcionOC DATE NOT NULL,
        CodigoMoneda NVARCHAR(3) NOT NULL DEFAULT N'PEN',
        Importe DECIMAL(19,2) NULL,
        Observacion NVARCHAR(2000) NULL,
        Estado NVARCHAR(30) NOT NULL DEFAULT N'REGISTRADA',
        CONSTRAINT FK_OC_Revision FOREIGN KEY (PresupuestoId, RevisionId) REFERENCES dbo.PresupuestoRevision(PresupuestoId, RevisionId),
        CONSTRAINT FK_OC_Moneda FOREIGN KEY (CodigoMoneda) REFERENCES dbo.Moneda(Codigo),
        CONSTRAINT CK_OC_Estado CHECK (Estado IN (N'REGISTRADA', N'TRABAJO_COMPLETADO', N'CERRADA_OPERATIVAMENTE', N'CANCELADA')),
        FechaRegistro DATETIME2(3) NOT NULL CONSTRAINT DF_OrdenCompraCliente_Registro DEFAULT SYSUTCDATETIME(),
        UsuarioRegistro NVARCHAR(128) NOT NULL,
        FechaActualizacion DATETIME2(3) NULL,
        UsuarioActualizacion NVARCHAR(128) NULL,
        VersionFila ROWVERSION NOT NULL,
        CONSTRAINT FK_OrdenCompraCliente_Autor FOREIGN KEY (UsuarioRegistro) REFERENCES dbo.AspNetUsers(Id),
        CONSTRAINT FK_OrdenCompraCliente_Editor FOREIGN KEY (UsuarioActualizacion) REFERENCES dbo.AspNetUsers(Id)
    );
END;
GO
