-- =============================================
-- Author:        FRANCO LARA
-- Create date:   22/06/2026
-- Description:   Agrega FechaEmision en CON_Asiento y despliega procedimientos de eliminacion para compras, ventas y asientos manuales.
-- =============================================

IF COL_LENGTH(N'dbo.CON_Asiento', N'FechaEmision') IS NULL
BEGIN
    ALTER TABLE dbo.CON_Asiento
        ADD FechaEmision DATE NULL;

    UPDATE dbo.CON_Asiento
    SET FechaEmision = FechaAsiento
    WHERE FechaEmision IS NULL;

    ALTER TABLE dbo.CON_Asiento
        ALTER COLUMN FechaEmision DATE NOT NULL;
END;

PRINT N'Desplegar adicionalmente los objetos actualizados desde StoreProcedure:';
PRINT N' - usp_CON_ListarAsientosPorEmpresa.sql';
PRINT N' - usp_CON_ObtenerAsiento.sql';
PRINT N' - usp_CON_GuardarAsientoManual.sql';
PRINT N' - usp_COM_GuardarCompraConAsiento.sql';
PRINT N' - usp_VEN_GuardarVentaConAsiento.sql';
PRINT N' - usp_COM_EliminarCompra.sql';
PRINT N' - usp_VEN_EliminarVenta.sql';
PRINT N' - usp_CON_EliminarAsiento.sql';
