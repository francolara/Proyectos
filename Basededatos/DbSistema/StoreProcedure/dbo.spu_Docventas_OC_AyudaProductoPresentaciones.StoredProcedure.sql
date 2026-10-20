-- =============================================
-- Author:        FRANCO LARA / Codex
-- Create date:   17/10/2026
-- Description:   Lista presentaciones y ultimo precio del proveedor para la ayuda de productos OC.
-- =============================================
CREATE OR ALTER PROCEDURE dbo.spu_Docventas_OC_AyudaProductoPresentaciones
    @IdEmpresa CHAR(2),
    @IdProveedor VARCHAR(20),
    @IdProducto VARCHAR(20),
    @IdLista VARCHAR(20)
AS
BEGIN
    SET NOCOUNT ON;

    BEGIN TRY
        SELECT
            PV.IdUM,
            UM.AbreUM AS GlsUM,
            CAST(PR.Factor AS DECIMAL(14, 2)) AS Factor,
            PP.ValorVenta AS VVUnit,
            PP.IGVVenta AS IGVUnit,
            PP.PrecioVenta AS PVUnit
        FROM dbo.PreciosVenta AS PV
        INNER JOIN dbo.UnidadMedida AS UM
            ON UM.IdUM = PV.IdUM
        INNER JOIN dbo.Presentaciones AS PR
            ON PR.IdEmpresa = PV.IdEmpresa
           AND PR.IdProducto = PV.IdProducto
           AND PR.IdUM = PV.IdUM
        CROSS APPLY (
            SELECT TOP (1)
                PPA.ValorVenta,
                PPA.IGVVenta,
                PPA.PrecioVenta
            FROM dbo.ProductosProveedor AS PPA
            WHERE PPA.IdEmpresa = PV.IdEmpresa
              AND PPA.IdProducto = PV.IdProducto
              AND PPA.IdProveedor = @IdProveedor
            ORDER BY PPA.Fecha DESC
        ) AS PP
        WHERE PV.IdEmpresa = @IdEmpresa
          AND PV.IdProducto = @IdProducto
          AND PV.IdLista = @IdLista
        ORDER BY PR.Factor;
    END TRY
    BEGIN CATCH
        DECLARE @ErrorMessage NVARCHAR(4000)
        DECLARE @ErrorSeverity INT
        DECLARE @ErrorState INT

        SELECT
            @ErrorMessage = ERROR_MESSAGE(),
            @ErrorSeverity = ERROR_SEVERITY(),
            @ErrorState = ERROR_STATE()

        RAISERROR (@ErrorMessage, @ErrorSeverity, @ErrorState)
    END CATCH
END
