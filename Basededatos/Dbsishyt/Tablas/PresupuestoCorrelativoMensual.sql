-- Firma: FRANCO LARA | Create date: 06/11/2026 | Numeración visible del presupuesto por año/mes de registro en Lima y correlativo mensual.
USE [SistemaWebHYT];
GO
SET NOCOUNT ON;
IF OBJECT_ID(N'dbo.PresupuestoCorrelativoMensual',N'U') IS NULL
BEGIN
    CREATE TABLE dbo.PresupuestoCorrelativoMensual
    (
        Periodo CHAR(6) NOT NULL CONSTRAINT PK_PresupuestoCorrelativoMensual PRIMARY KEY,
        Ultimo BIGINT NOT NULL CONSTRAINT CK_PresupuestoCorrelativoMensual_Ultimo CHECK(Ultimo>=0)
    );
END;
GO
