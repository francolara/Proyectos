-- =============================================
-- Author:        FRANCO LARA
-- Create date:   09/11/2026
-- Description:   PlanPagoDetalleHYT
-- =============================================
USE [SistemaWebHYT];
GO
IF OBJECT_ID(N'dbo.PlanPagoDetalleHYT', N'U') IS NULL
BEGIN
    CREATE TABLE dbo.PlanPagoDetalleHYT
    (
        PlanId BIGINT NOT NULL CONSTRAINT FK_PlanPagoDetalleHYT_Plan REFERENCES dbo.PlanPagoHYT(PlanId),
        TrabajadorId BIGINT NOT NULL CONSTRAINT FK_PlanPagoDetalleHYT_Trabajador REFERENCES dbo.TrabajadorHYT(TrabajadorId),
        Nombre NVARCHAR(180) NOT NULL,
        Tipo VARCHAR(10) NOT NULL,
        Base DECIMAL(18,2) NOT NULL,
        MinutosExtra INT NOT NULL,
        ImporteExtra DECIMAL(18,2) NOT NULL,
        Descuento DECIMAL(18,2) NOT NULL,
        CONSTRAINT PK_PlanPagoDetalleHYT PRIMARY KEY (PlanId, TrabajadorId),
        CONSTRAINT CK_PlanPagoDetalleHYT_Importes CHECK (Base >= 0 AND MinutosExtra >= 0
            AND ImporteExtra >= 0 AND Descuento >= 0 AND Descuento <= Base + ImporteExtra)
    );
END;
GO
