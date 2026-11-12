-- Firma: FRANCO LARA | Create date: 11/11/2026 | Permite configurar una cuenta de cargo en soles o dólares.
USE [SistemaWebHYT];
GO
SET NOCOUNT ON;
SET XACT_ABORT ON;
IF OBJECT_ID(N'dbo.CK_TelecreditoHYT_CuentaMoneda',N'C') IS NULL
BEGIN
    BEGIN TRANSACTION;
    DECLARE @Eliminar NVARCHAR(MAX);
    SELECT @Eliminar=STRING_AGG(CONVERT(NVARCHAR(MAX),
        N'ALTER TABLE dbo.TelecreditoHYT DROP CONSTRAINT '+QUOTENAME(regla.name)+N';'),N' ')
    FROM sys.check_constraints AS regla
    WHERE regla.parent_object_id=OBJECT_ID(N'dbo.TelecreditoHYT')
        AND COL_NAME(regla.parent_object_id,regla.parent_column_id) IN (N'NumeroCuenta',N'Moneda');
    IF @Eliminar IS NOT NULL EXEC sys.sp_executesql @Eliminar;

    ALTER TABLE dbo.TelecreditoHYT WITH CHECK ADD CONSTRAINT CK_TelecreditoHYT_CuentaMoneda CHECK (
        Moneda IN ('PEN','USD') AND DATALENGTH(NumeroCuenta)=13
        AND NumeroCuenta COLLATE Latin1_General_100_BIN2 NOT LIKE '%[^0-9]%'
        AND SUBSTRING(NumeroCuenta,11,1)=CASE Moneda WHEN 'PEN' THEN '0' WHEN 'USD' THEN '1' END);
    COMMIT TRANSACTION;
END;
GO
