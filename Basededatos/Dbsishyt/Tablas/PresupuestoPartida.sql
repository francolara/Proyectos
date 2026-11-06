-- =============================================
-- Author:        FRANCO LARA
-- Create date:   03/11/2026
-- Description:   Define PresupuestoPartida para la base inicial de HYT.
-- =============================================
-- Firma: FRANCO LARA | Create date: 06/11/2026 | Tipo de servicio admite el nombre completo del maestro.
USE [SistemaWebHYT];
GO
SET NOCOUNT ON;
IF OBJECT_ID(N'dbo.PresupuestoPartida', N'U') IS NULL
BEGIN
    CREATE TABLE dbo.PresupuestoPartida
    (
        PartidaId BIGINT IDENTITY NOT NULL CONSTRAINT PK_PresupuestoPartida PRIMARY KEY,
        RevisionId BIGINT NOT NULL,
        NumeroLinea INT NOT NULL,
        Grupo NVARCHAR(200) NULL,
        MaterialId BIGINT NULL,
        PartidaTarifarioId BIGINT NULL,
        Posicion NVARCHAR(60) NULL,
        CodigoInterno NVARCHAR(60) NULL,
        CodigoMaterialSAP NVARCHAR(60) NULL,
        CentroSAPFICO NVARCHAR(50) NULL,
        NumeroNecesidad NVARCHAR(60) NULL,
        TipoServicio NVARCHAR(150) NULL,
        Descripcion NVARCHAR(1500) NOT NULL,
        Unidad NVARCHAR(30) NOT NULL,
        Cantidad DECIMAL(19,6) NOT NULL,
        PrecioOrigen DECIMAL(19,6) NULL,
        CodigoMonedaOrigen NVARCHAR(3) NULL,
        CodigoTipoIGVOrigen NVARCHAR(10) NULL,
        PrecioUnitario DECIMAL(19,6) NULL,
        BaseImponible DECIMAL(19,2) NULL,
        Impuesto DECIMAL(19,2) NULL,
        Total DECIMAL(19,2) NULL,
        CONSTRAINT UQ_PresupuestoPartida_Linea UNIQUE (RevisionId, NumeroLinea),
        CONSTRAINT FK_Partida_Revision FOREIGN KEY (RevisionId) REFERENCES dbo.PresupuestoRevision(RevisionId),
        CONSTRAINT FK_Partida_Material FOREIGN KEY (MaterialId) REFERENCES dbo.ServicioMaterial(MaterialId),
        CONSTRAINT FK_Partida_Tarifa FOREIGN KEY (PartidaTarifarioId) REFERENCES dbo.TarifarioPartida(PartidaId),
        CONSTRAINT FK_Partida_Moneda FOREIGN KEY (CodigoMonedaOrigen) REFERENCES dbo.Moneda(Codigo),
        CONSTRAINT FK_Partida_IGV FOREIGN KEY (CodigoTipoIGVOrigen) REFERENCES dbo.TipoIGV(Codigo),
        CONSTRAINT CK_Partida_Cantidad CHECK (Cantidad > 0),
        CONSTRAINT CK_Partida_Precio CHECK (PrecioUnitario IS NULL OR PrecioUnitario >= 0)
    );
END;
GO
