-- =============================================
-- Author:        FRANCO LARA
-- Create date:   16/10/2026
-- Description:   Lista productos para la ayuda de ventas, con stock y separación por sucursal y almacén.
-- =============================================
CREATE OR ALTER PROCEDURE dbo.spu_Docventa_Lista_Productos
    @IdEmpresa CHAR(2),
    @IdLista CHAR(8),
    @IdAlmacen CHAR(8),
    @IdTipoProducto CHAR(6),
    @IdNivel CHAR(8),
    @Busqueda VARCHAR(300),
    @IdSucursal CHAR(8) = ''
AS
BEGIN
    SET NOCOUNT ON;

    BEGIN TRY
        DECLARE @TextoBusqueda VARCHAR(300) = LTRIM(RTRIM(ISNULL(@Busqueda, '')));
        DECLARE @PatronBusqueda VARCHAR(302);

        SET @TextoBusqueda = REPLACE(@TextoBusqueda, '%', '');
        SET @PatronBusqueda = '%' + @TextoBusqueda + '%';

        IF @IdTipoProducto <> '06002'
        BEGIN
            ;WITH Movimientos AS
            (
                SELECT
                    VD.IdEmpresa,
                    VD.IdProducto,
                    SUM(CASE WHEN VD.TipoVale = 'I' THEN VD.Cantidad ELSE VD.Cantidad * -1 END) AS CantidadStock
                FROM dbo.ValesCab AS VC
                INNER JOIN dbo.PeriodosINV AS PI
                    ON PI.IdEmpresa = VC.IdEmpresa
                   AND PI.IdPeriodoINV = VC.IdPeriodoINV
                   AND PI.IdSucursal = VC.IdSucursal
                INNER JOIN dbo.ValesDet AS VD
                    ON VD.IdEmpresa = VC.IdEmpresa
                   AND VD.IdSucursal = VC.IdSucursal
                   AND VD.TipoVale = VC.TipoVale
                   AND VD.IdValesCab = VC.IdValesCab
                WHERE VC.IdEmpresa = @IdEmpresa
                  AND VC.EstValeCab <> 'ANU'
                  AND PI.EstPeriodoInv = 'ACT'
                  AND (@IdAlmacen = '' OR VC.IdAlmacen = @IdAlmacen)
                  AND (@IdSucursal = '' OR VC.IdSucursal = @IdSucursal)
                GROUP BY VD.IdEmpresa, VD.IdProducto
            ),
            Separaciones AS
            (
                SELECT
                    PA.IdEmpresa,
                    PA.IdProducto,
                    PA.IdUMCompra,
                    SUM(CASE WHEN ISNULL(PA.Separacion, 0) < 0 THEN 0 ELSE ISNULL(PA.Separacion, 0) END) AS Separacion
                FROM dbo.ProductosAlmacen AS PA
                WHERE PA.IdEmpresa = @IdEmpresa
                  AND (@IdAlmacen = '' OR PA.IdAlmacen = @IdAlmacen)
                  AND (@IdSucursal = '' OR PA.IdSucursal = @IdSucursal)
                GROUP BY PA.IdEmpresa, PA.IdProducto, PA.IdUMCompra
            )
            SELECT
                P.CodigoRapido,
                P.IdProducto,
                P.GlsProducto,
                M.GlsMarca,
                P.IdUMCompra AS IdUMVenta,
                U.GlsUM,
                O.IdMoneda AS GlsMoneda,
                CASE WHEN P.AfectoIGV = 1 THEN 'S' ELSE 'N' END AS Afecto,
                CAST(ISNULL(MV.CantidadStock, 0) AS NUMERIC(12, 2)) AS Stock,
                TP.GlsTallaPeso,
                P.IdFabricante,
                CAST(ISNULL(SP.Separacion, 0) AS NUMERIC(12, 2)) AS Separacion,
                CAST(ISNULL(MV.CantidadStock, 0) - ISNULL(SP.Separacion, 0) AS NUMERIC(12, 2)) AS Disponible
            FROM dbo.Productos AS P
            INNER JOIN dbo.Marcas AS M
                ON M.IdEmpresa = P.IdEmpresa
               AND M.IdMarca = P.IdMarca
            INNER JOIN dbo.UnidadMedida AS U
                ON U.IdUM = P.IdUMCompra
            INNER JOIN dbo.Monedas AS O
                ON O.IdMoneda = P.IdMoneda
            LEFT JOIN dbo.TallaPeso AS TP
                ON TP.IdEmpresa = P.IdEmpresa
               AND TP.IdTallaPeso = P.IdTallaPeso
            LEFT JOIN Movimientos AS MV
                ON MV.IdEmpresa = P.IdEmpresa
               AND MV.IdProducto = P.IdProducto
            LEFT JOIN Separaciones AS SP
                ON SP.IdEmpresa = P.IdEmpresa
               AND SP.IdProducto = P.IdProducto
               AND SP.IdUMCompra = P.IdUMCompra
            WHERE P.IdEmpresa = @IdEmpresa
              AND P.EstProducto = 'A'
              AND P.IdTipoProducto = @IdTipoProducto
              AND (@IdNivel = '' OR P.IdNivel = @IdNivel)
              AND EXISTS
              (
                  SELECT 1
                  FROM dbo.PreciosVenta AS PV
                  WHERE PV.IdEmpresa = P.IdEmpresa
                    AND PV.IdLista = @IdLista
                    AND PV.IdProducto = P.IdProducto
              )
              AND
              (
                  @TextoBusqueda = ''
                  OR P.IdProducto LIKE @PatronBusqueda
                  OR P.GlsProducto LIKE @PatronBusqueda
                  OR P.CodigoRapido LIKE @PatronBusqueda
                  OR P.IdFabricante LIKE @PatronBusqueda
              )
            ORDER BY P.IdProducto;
        END
        ELSE
        BEGIN
            SELECT
                CAST('' AS VARCHAR(50)) AS CodigoRapido,
                P.IdProducto,
                P.GlsProducto,
                CAST('' AS VARCHAR(185)) AS GlsMarca,
                P.IdUMVenta,
                U.GlsUM,
                O.IdMoneda AS GlsMoneda,
                CASE WHEN P.AfectoIGV = 1 THEN 'S' ELSE 'N' END AS Afecto,
                CAST(0 AS NUMERIC(12, 2)) AS Stock,
                CAST('' AS VARCHAR(185)) AS GlsTallaPeso,
                P.IdFabricante,
                CAST(0 AS NUMERIC(12, 2)) AS Separacion,
                CAST(0 AS NUMERIC(12, 2)) AS Disponible
            FROM dbo.Productos AS P
            INNER JOIN dbo.Monedas AS O
                ON O.IdMoneda = P.IdMoneda
            INNER JOIN dbo.UnidadMedida AS U
                ON U.IdUM = P.IdUMVenta
            WHERE P.IdEmpresa = @IdEmpresa
              AND P.IdTipoProducto = @IdTipoProducto
              AND P.EstProducto = 'A'
              AND (@IdNivel = '' OR P.IdNivel = @IdNivel)
              AND
              (
                  @TextoBusqueda = ''
                  OR P.IdProducto LIKE @PatronBusqueda
                  OR P.GlsProducto LIKE @PatronBusqueda
                  OR P.CodigoRapido LIKE @PatronBusqueda
                  OR P.IdFabricante LIKE @PatronBusqueda
              )
            ORDER BY P.IdProducto;
        END
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
