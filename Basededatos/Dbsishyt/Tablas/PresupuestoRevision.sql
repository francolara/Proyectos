-- =============================================
-- Author:        FRANCO LARA
-- Create date:   03/11/2026
-- Description:   Define PresupuestoRevision para la base inicial de HYT.
-- =============================================
USE [SistemaWebHYT];
GO
SET NOCOUNT ON;
IF OBJECT_ID(N'dbo.PresupuestoRevision', N'U') IS NULL
BEGIN
    CREATE TABLE dbo.PresupuestoRevision
    (
        RevisionId BIGINT IDENTITY NOT NULL CONSTRAINT PK_PresupuestoRevision PRIMARY KEY,
        PresupuestoId BIGINT NOT NULL,
        NumeroRevision INT NOT NULL,
        FechaEmision DATE NOT NULL,
        NumeroCotizacionCliente NVARCHAR(60) NULL,
        ClienteNombre NVARCHAR(200) NOT NULL,
        ClienteDocumento NVARCHAR(30) NULL,
        ClienteDireccion NVARCHAR(300) NULL,
        TiendaNombre NVARCHAR(200) NOT NULL,
        TiendaDireccion NVARCHAR(300) NULL,
        FechaEjecucion DATE NULL,
        Trimestre NVARCHAR(40) NULL,
        VigenciaTexto NVARCHAR(100) NULL,
        CodigoSAPProveedor NVARCHAR(60) NULL,
        ContratoMarco NVARCHAR(60) NULL,
        Especialidad NVARCHAR(150) NULL,
        Coordinador NVARCHAR(150) NULL,
        Encargado NVARCHAR(150) NULL,
        Consideraciones NVARCHAR(MAX) NULL,
        Firmante NVARCHAR(150) NULL,
        CargoFirmante NVARCHAR(150) NULL,
        CodigoMoneda NVARCHAR(3) NOT NULL DEFAULT N'PEN',
        CodigoTipoIGV NVARCHAR(10) NOT NULL,
        FechaDesdeTasaIGV DATE NOT NULL,
        TasaIGVAplicada DECIMAL(9,6) NOT NULL,
        Estado NVARCHAR(20) NOT NULL DEFAULT N'BORRADOR',
        Subtotal DECIMAL(19,2) NULL,
        Impuesto DECIMAL(19,2) NULL,
        Total DECIMAL(19,2) NULL,
        CONSTRAINT UQ_PresupuestoRevision_Numero UNIQUE (PresupuestoId, NumeroRevision),
        CONSTRAINT UQ_PresupuestoRevision_Vinculo UNIQUE (PresupuestoId, RevisionId),
        CONSTRAINT FK_Revision_Presupuesto FOREIGN KEY (PresupuestoId) REFERENCES dbo.Presupuesto(PresupuestoId),
        CONSTRAINT FK_Revision_Moneda FOREIGN KEY (CodigoMoneda) REFERENCES dbo.Moneda(Codigo),
        CONSTRAINT FK_Revision_TipoIGV FOREIGN KEY (CodigoTipoIGV) REFERENCES dbo.TipoIGV(Codigo),
        CONSTRAINT FK_Revision_Tasa FOREIGN KEY (FechaDesdeTasaIGV, TasaIGVAplicada) REFERENCES dbo.TasaIGV(FechaDesde, Valor),
        CONSTRAINT CK_Revision_Tasa CHECK (TasaIGVAplicada >= 0 AND TasaIGVAplicada <= 1),
        CONSTRAINT CK_Revision_Estado CHECK (Estado IN (N'BORRADOR', N'ENVIADO', N'OBSERVADO', N'APROBADO', N'RECHAZADO', N'VENCIDO', N'CANCELADO')),
        FechaRegistro DATETIME2(3) NOT NULL CONSTRAINT DF_PresupuestoRevision_Registro DEFAULT SYSUTCDATETIME(),
        UsuarioRegistro NVARCHAR(128) NOT NULL,
        FechaActualizacion DATETIME2(3) NULL,
        UsuarioActualizacion NVARCHAR(128) NULL,
        VersionFila ROWVERSION NOT NULL,
        CONSTRAINT FK_PresupuestoRevision_Autor FOREIGN KEY (UsuarioRegistro) REFERENCES dbo.AspNetUsers(Id),
        CONSTRAINT FK_PresupuestoRevision_Editor FOREIGN KEY (UsuarioActualizacion) REFERENCES dbo.AspNetUsers(Id)
    );
END;
GO
