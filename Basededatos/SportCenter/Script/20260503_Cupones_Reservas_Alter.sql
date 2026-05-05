USE [DbSportCenter]
GO
-- Firma: FRANCO LARA
-- Create date: 03/05/2026
-- Descripcion: agrega columnas para trazabilidad de cupon aplicado en reservas.
IF COL_LENGTH('dbo.Reservas', 'CodigoCuponAplicado') IS NULL
BEGIN
    ALTER TABLE dbo.Reservas ADD CodigoCuponAplicado NVARCHAR(30) NULL;
END
GO
IF COL_LENGTH('dbo.Reservas', 'DescuentoCupon') IS NULL
BEGIN
    ALTER TABLE dbo.Reservas ADD DescuentoCupon DECIMAL(10,2) NULL CONSTRAINT DF_Reservas_DescuentoCupon DEFAULT ((0));
END
GO
