
GO
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO

-- =============================================
-- Author:      FRANCO LARA / Codex
-- Create date: 13/10/2026
-- Description: Obtiene el producto activo y su precio de lista para el detalle de ventas.
-- =============================================
CREATE OR ALTER PROCEDURE dbo.spu_Docventas_ObtieneProductoPrecio
    @IdEmpresa CHAR(2),
    @IdLista CHAR(8),
    @CodigoBusqueda VARCHAR(40)
AS
BEGIN
    SET NOCOUNT ON;

    BEGIN TRY
        ;WITH ProductoBuscado AS
        (
            SELECT TOP (1)
                p.IdProducto,
                p.GlsProducto,
                p.IdUMVenta,
                p.IdFabricante,
                p.IdMarca,
                p.AfectoIGV,
                p.IdTipoProducto,
                p.CodigoRapido,
                p.AfectoDctoEspecial,
                p.IndAfectoIVAP
            FROM dbo.Productos AS p
            WHERE p.IdEmpresa = @IdEmpresa
              AND p.EstProducto = 'A'
              AND (
                    p.IdProducto = @CodigoBusqueda
                    OR p.IdFabricante = @CodigoBusqueda
                    OR p.CodigoRapido = @CodigoBusqueda
                  )
            ORDER BY
                CASE
                    WHEN p.IdProducto = @CodigoBusqueda THEN 1
                    WHEN p.IdFabricante = @CodigoBusqueda THEN 2
                    ELSE 3
                END,
                p.IdProducto
        )
        SELECT
            p.IdProducto,
            p.GlsProducto,
            p.IdUMVenta AS IdUM,
            p.CodigoRapido,
            p.IdFabricante,
            p.IdMarca,
            ISNULL(m.GlsMarca, '') AS GlsMarca,
            CAST(p.AfectoIGV AS INT) AS AfectoIGV,
            p.IdTipoProducto,
            ISNULL(lp.IdMoneda, '') AS IdMoneda,
            CASE WHEN p.IdTipoProducto = '06002' THEN '' ELSE ISNULL(um.AbreUM, '') END AS GlsUM,
            ISNULL(pv.VVUnit, 0) AS VVUnit,
            ISNULL(md.MaxDcto, 0) AS MaxDcto,
            CASE WHEN p.IdTipoProducto = '06002' THEN CAST(1 AS DECIMAL(14, 5)) ELSE ISNULL(pr.Factor, 1) END AS Factor,
            CAST(p.AfectoDctoEspecial AS INT) AS AfectoDctoEspecial,
            ISNULL(p.IndAfectoIVAP, '') AS IndAfectoIVAP
        FROM ProductoBuscado AS p
        LEFT JOIN dbo.Marcas AS m
            ON m.IdEmpresa = @IdEmpresa
           AND m.IdMarca = p.IdMarca
        LEFT JOIN dbo.ListaPrecios AS lp
            ON lp.IdEmpresa = @IdEmpresa
           AND lp.IdLista = @IdLista
        OUTER APPLY
        (
            SELECT TOP (1)
                v.IdUM,
                v.VVUnit
            FROM dbo.PreciosVenta AS v
            WHERE v.IdEmpresa = @IdEmpresa
              AND v.IdLista = @IdLista
              AND v.IdProducto = p.IdProducto
              AND (p.IdTipoProducto = '06002' OR v.IdUM = p.IdUMVenta)
            ORDER BY
                CASE WHEN v.IdUM = p.IdUMVenta THEN 0 ELSE 1 END,
                v.IdUM
        ) AS pv
        OUTER APPLY
        (
            SELECT v.MaxDcto
            FROM dbo.PreciosVenta AS v
            WHERE v.IdEmpresa = @IdEmpresa
              AND v.IdLista = @IdLista
              AND v.IdProducto = p.IdProducto
              AND v.IdUM = p.IdUMVenta
        ) AS md
        LEFT JOIN dbo.Presentaciones AS pr
            ON pr.IdEmpresa = @IdEmpresa
           AND pr.IdProducto = p.IdProducto
           AND pr.IdUM = pv.IdUM
           AND p.IdTipoProducto <> '06002'
        LEFT JOIN dbo.UnidadMedida AS um
            ON um.IdUM = pv.IdUM
           AND p.IdTipoProducto <> '06002';
    END TRY
    BEGIN CATCH
        DECLARE @ErrorMessage NVARCHAR(4000);
        DECLARE @ErrorSeverity INT;
        DECLARE @ErrorState INT;

        SELECT
            @ErrorMessage = ERROR_MESSAGE(),
            @ErrorSeverity = ERROR_SEVERITY(),
            @ErrorState = ERROR_STATE();

        RAISERROR (@ErrorMessage, @ErrorSeverity, @ErrorState);
    END CATCH
END
GO
