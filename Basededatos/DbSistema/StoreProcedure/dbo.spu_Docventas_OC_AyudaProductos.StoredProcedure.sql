-- =============================================
-- Author:        FRANCO LARA / Codex
-- Create date:   17/10/2026
-- Description:   Lista productos para la ayuda propia de ordenes de compra.
-- =============================================
CREATE OR ALTER PROCEDURE dbo.spu_Docventas_OC_AyudaProductos
    @IdEmpresa CHAR(2),
    @IdSucursal VARCHAR(8),
    @IdAlmacen VARCHAR(8),
    @IdProveedor VARCHAR(20),
    @Busqueda VARCHAR(300),
    @IdNivel VARCHAR(20),
    @IdTipoProducto CHAR(5),
    @IndPedido BIT,
    @IndValidaStock BIT,
    @IndUMVenta BIT,
    @ValidaProductoProveedor CHAR(1)
AS
BEGIN
    SET NOCOUNT ON;

    BEGIN TRY
        IF @IdTipoProducto = '06002'
        BEGIN
            SELECT
                '0' AS CHK,
                P.IdProducto,
                P.CodigoRapido,
                P.GlsProducto,
                ISNULL(M.GlsMarca, '') AS GlsMarca,
                ISNULL(U.GlsUM, '') AS GlsUM,
                O.GlsMoneda,
                CAST('' AS VARCHAR(100)) AS GlsTallaPeso,
                CASE WHEN P.AfectoIGV = 1 THEN 'S' ELSE 'N' END AS Afecto,
                CASE WHEN @IndUMVenta = 1 THEN P.IdUMVenta ELSE P.IdUMCompra END AS IdUMVenta,
                CAST(0 AS DECIMAL(12, 2)) AS Stock,
                P.IdFabricante,
                CAST('' AS VARCHAR(200)) AS GlsDocReferencia
            FROM dbo.Productos AS P
            INNER JOIN dbo.Monedas AS O
                ON O.IdMoneda = P.IdMoneda
            LEFT JOIN dbo.Marcas AS M
                ON M.IdEmpresa = P.IdEmpresa
               AND M.IdMarca = P.IdMarca
            LEFT JOIN dbo.UnidadMedida AS U
                ON U.IdUM = CASE WHEN @IndUMVenta = 1 THEN P.IdUMVenta ELSE P.IdUMCompra END
            WHERE P.IdEmpresa = @IdEmpresa
              AND P.IdTipoProducto = @IdTipoProducto
              AND (
                    @Busqueda = ''
                    OR P.GlsProducto LIKE '%' + @Busqueda + '%'
                    OR P.CodigoRapido LIKE '%' + @Busqueda + '%'
                    OR P.IdFabricante LIKE '%' + @Busqueda + '%'
                  )
              AND (@IdNivel = '' OR P.IdNivel = @IdNivel)
            ORDER BY P.GlsProducto, P.IdProducto;

            RETURN;
        END

        SELECT
            '0' AS CHK,
            P.IdProducto,
            P.CodigoRapido,
            P.GlsProducto,
            M.GlsMarca,
            U.GlsUM,
            O.GlsMoneda,
            T.GlsTallaPeso,
            CASE WHEN P.AfectoIGV = 1 THEN 'S' ELSE 'N' END AS Afecto,
            CASE WHEN @IndUMVenta = 1 THEN P.IdUMVenta ELSE P.IdUMCompra END AS IdUMVenta,
            CAST(
                CASE
                    WHEN @IdTipoProducto = '06002' OR @IndPedido = 1 THEN 0
                    WHEN @IndUMVenta = 1 THEN ISNULL(PA.CantidadStock / NULLIF(PR.Factor, 0), 0)
                    ELSE ISNULL(PA.CantidadStock, 0)
                END
                AS DECIMAL(12, 2)
            ) AS Stock,
            P.IdFabricante,
            CAST('' AS VARCHAR(200)) AS GlsDocReferencia
        FROM dbo.Productos AS P
        INNER JOIN dbo.Marcas AS M
            ON M.IdEmpresa = P.IdEmpresa
           AND M.IdMarca = P.IdMarca
        INNER JOIN dbo.UnidadMedida AS U
            ON U.IdUM = CASE WHEN @IndUMVenta = 1 THEN P.IdUMVenta ELSE P.IdUMCompra END
        INNER JOIN dbo.Monedas AS O
            ON O.IdMoneda = P.IdMoneda
        LEFT JOIN dbo.ProductosAlmacen AS PA
            ON PA.IdEmpresa = P.IdEmpresa
           AND PA.IdSucursal = @IdSucursal
           AND PA.IdAlmacen = @IdAlmacen
           AND PA.IdProducto = P.IdProducto
           AND PA.IdUMCompra = P.IdUMCompra
        LEFT JOIN dbo.Presentaciones AS PR
            ON PR.IdEmpresa = P.IdEmpresa
           AND PR.IdProducto = P.IdProducto
           AND PR.IdUM = P.IdUMVenta
           AND @IndUMVenta = 1
        LEFT JOIN dbo.TallaPeso AS T
            ON T.IdEmpresa = P.IdEmpresa
           AND T.IdTallaPeso = P.IdTallaPeso
        WHERE P.IdEmpresa = @IdEmpresa
          AND P.IdTipoProducto = @IdTipoProducto
          AND (
                @Busqueda = ''
                OR P.GlsProducto LIKE '%' + @Busqueda + '%'
                OR P.CodigoRapido LIKE '%' + @Busqueda + '%'
                OR P.IdFabricante LIKE '%' + @Busqueda + '%'
              )
          AND (@IdNivel = '' OR P.IdNivel = @IdNivel)
          AND (
                @IndPedido = 1
                OR PA.IdProducto IS NOT NULL
              )
          AND (
                @IndPedido = 1
                OR @IndValidaStock = 0
                OR CASE
                       WHEN @IndUMVenta = 1 THEN ISNULL(PA.CantidadStock / NULLIF(PR.Factor, 0), 0)
                       ELSE ISNULL(PA.CantidadStock, 0)
                   END > 0
              )
          AND (@IndPedido = 1 OR @IndUMVenta = 0 OR PR.IdProducto IS NOT NULL)
          AND (
                @IndPedido = 0
                OR @ValidaProductoProveedor = 'S'
                OR EXISTS (
                    SELECT 1
                    FROM dbo.ProductosProveedor AS PP
                    WHERE PP.IdEmpresa = P.IdEmpresa
                      AND PP.IdProducto = P.IdProducto
                      AND PP.IdProveedor = @IdProveedor
                )
              )
        ORDER BY P.GlsProducto, P.IdProducto;
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
