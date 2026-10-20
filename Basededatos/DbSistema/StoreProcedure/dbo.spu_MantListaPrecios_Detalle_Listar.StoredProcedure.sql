-- =============================================
-- Author:      FRANCO LARA / Codex
-- Create date: 15/10/2026
-- Description: Lista el detalle de una lista de precios con niveles dinamicos y stock opcional.
-- =============================================
CREATE OR ALTER PROCEDURE dbo.spu_MantListaPrecios_Detalle_Listar
    @IdEmpresa CHAR(2),
    @IdLista CHAR(8),
    @IdNivel CHAR(8) = '',
    @Busqueda VARCHAR(255) = '',
    @VisualizaAlmacenes BIT = 0
AS
BEGIN
    SET NOCOUNT ON;

    BEGIN TRY
        DECLARE @Sql NVARCHAR(MAX) = N'';
        DECLARE @ColumnasNivel NVARCHAR(MAX) = N'';
        DECLARE @UnionesNivel NVARCHAR(MAX) = N'';
        DECLARE @ColumnasAlmacen NVARCHAR(MAX) = N'';
        DECLARE @UnionesAlmacen NVARCHAR(MAX) = N'';
        DECLARE @AliasActual SYSNAME;
        DECLARE @AliasAnterior SYSNAME = N'';
        DECLARE @Indice INT = 0;
        DECLARE @Peso INT;
        DECLARE @BusquedaLike VARCHAR(257);

        SET @Busqueda = LTRIM(RTRIM(ISNULL(@Busqueda, '')));
        SET @BusquedaLike = '%' + @Busqueda + '%';

        DECLARE CursorNiveles CURSOR LOCAL FAST_FORWARD FOR
            SELECT TN.Peso
            FROM dbo.TiposNiveles AS TN
            WHERE TN.IdEmpresa = @IdEmpresa
            ORDER BY TN.Peso ASC;

        OPEN CursorNiveles;
        FETCH NEXT FROM CursorNiveles INTO @Peso;

        WHILE @@FETCH_STATUS = 0
        BEGIN
            SET @Indice = @Indice + 1;
            SET @AliasActual = N'N' + RIGHT(N'00' + CONVERT(NVARCHAR(10), @Indice), 2);

            SET @ColumnasNivel = @ColumnasNivel
                + @AliasActual + N'.IdNivel AS IdNivel' + RIGHT(N'00' + CONVERT(NVARCHAR(10), @Indice), 2) + N', '
                + @AliasActual + N'.GlsNivel AS GlsNivel' + RIGHT(N'00' + CONVERT(NVARCHAR(10), @Indice), 2) + N', ';

            IF @Indice = 1
                SET @UnionesNivel = @UnionesNivel
                    + N' INNER JOIN dbo.Niveles AS ' + @AliasActual
                    + N' ON P.IdNivel = ' + @AliasActual + N'.IdNivel'
                    + N' AND ' + @AliasActual + N'.IdEmpresa = @PIdEmpresa';
            ELSE
                SET @UnionesNivel = @UnionesNivel
                    + N' INNER JOIN dbo.Niveles AS ' + @AliasActual
                    + N' ON ' + @AliasAnterior + N'.IdNivelPred = ' + @AliasActual + N'.IdNivel'
                    + N' AND ' + @AliasActual + N'.IdEmpresa = @PIdEmpresa';

            SET @AliasAnterior = @AliasActual;
            FETCH NEXT FROM CursorNiveles INTO @Peso;
        END;

        CLOSE CursorNiveles;
        DEALLOCATE CursorNiveles;

        IF @VisualizaAlmacenes = 1
        BEGIN
            SET @ColumnasAlmacen = N', P01.StkPrincipal, P03.StkDansey';
            SET @UnionesAlmacen = N'
                LEFT JOIN (
                    SELECT
                        VD.IdProducto,
                        VD.IdUM,
                        VC.IdEmpresa,
                        VC.IdSucursal,
                        CAST(SUM(CASE WHEN VD.TipoVale = ''I'' THEN VD.Cantidad ELSE VD.Cantidad * -1 END) AS NUMERIC(12, 2)) AS StkPrincipal
                    FROM dbo.ValesDet AS VD
                    INNER JOIN dbo.ValesCab AS VC
                        ON VD.IdValesCab = VC.IdValesCab
                       AND VD.IdEmpresa = VC.IdEmpresa
                       AND VD.IdSucursal = VC.IdSucursal
                       AND VD.TipoVale = VC.TipoVale
                    INNER JOIN dbo.PeriodosINV AS PI
                        ON VC.IdEmpresa = PI.IdEmpresa
                       AND VC.IdPeriodoINV = PI.IdPeriodoINV
                       AND VC.IdSucursal = PI.IdSucursal
                    INNER JOIN dbo.Productos AS PS
                        ON VD.IdEmpresa = PS.IdEmpresa
                       AND VD.IdProducto = PS.IdProducto
                    WHERE VC.IdEmpresa = @PIdEmpresa
                      AND VC.IdSucursal = ''08090001''
                      AND VC.IdAlmacen = ''11010001''
                      AND VC.EstValeCab <> ''ANU''
                      AND PI.EstPeriodoInv = ''ACT''
                      AND PS.GlsProducto LIKE @PBusquedaLike
                    GROUP BY VD.IdProducto, VD.IdUM, VC.IdEmpresa, VC.IdSucursal
                ) AS P01
                    ON P.IdEmpresa = P01.IdEmpresa
                   AND P.IdProducto = P01.IdProducto
                   AND P.IdUMCompra = P01.IdUM
                LEFT JOIN (
                    SELECT
                        VD.IdProducto,
                        VD.IdUM,
                        VC.IdEmpresa,
                        VC.IdSucursal,
                        CAST(SUM(CASE WHEN VD.TipoVale = ''I'' THEN VD.Cantidad ELSE VD.Cantidad * -1 END) AS NUMERIC(12, 2)) AS StkDansey
                    FROM dbo.ValesDet AS VD
                    INNER JOIN dbo.ValesCab AS VC
                        ON VD.IdValesCab = VC.IdValesCab
                       AND VD.IdEmpresa = VC.IdEmpresa
                       AND VD.IdSucursal = VC.IdSucursal
                       AND VD.TipoVale = VC.TipoVale
                    INNER JOIN dbo.PeriodosINV AS PI
                        ON VC.IdEmpresa = PI.IdEmpresa
                       AND VC.IdPeriodoINV = PI.IdPeriodoINV
                       AND VC.IdSucursal = PI.IdSucursal
                    INNER JOIN dbo.Productos AS PS
                        ON VD.IdEmpresa = PS.IdEmpresa
                       AND VD.IdProducto = PS.IdProducto
                    WHERE VC.IdEmpresa = ''01''
                      AND VC.IdSucursal = ''08090004''
                      AND VC.IdAlmacen = ''11030001''
                      AND VC.EstValeCab <> ''ANU''
                      AND PI.EstPeriodoInv = ''ACT''
                      AND PS.GlsProducto LIKE @PBusquedaLike
                    GROUP BY VD.IdProducto, VD.IdUM, VC.IdEmpresa, VC.IdSucursal
                ) AS P03
                    ON P.IdEmpresa = P03.IdEmpresa
                   AND P.IdProducto = P03.IdProducto
                   AND P.IdUMCompra = P03.IdUM';
        END;

        SET @Sql = N'SELECT ' + @ColumnasNivel
            + N'(L.IdProducto + L.IdUM) AS Item, L.IdProducto, P.GlsProducto, P.IdFabricante, UVC.AbreUM AS UMCompra, '
            + N'P.AfectoIGV AS Afecto, L.IdUM, UVL.AbreUM AS GlsUM, PR.Factor, L.VVUnit, L.IGVUnit, L.PVUnit, '
            + N'L.CostoUnit, L.FactorUnit, L.Factor2Unit'
            + @ColumnasAlmacen
            + N' FROM dbo.PreciosVenta AS L'
            + N' INNER JOIN dbo.Productos AS P ON L.IdEmpresa = P.IdEmpresa AND L.IdProducto = P.IdProducto'
            + N' INNER JOIN dbo.UnidadMedida AS UVC ON P.IdUMCompra = UVC.IdUM'
            + N' INNER JOIN dbo.UnidadMedida AS UVL ON L.IdUM = UVL.IdUM'
            + N' INNER JOIN dbo.Presentaciones AS PR ON L.IdProducto = PR.IdProducto AND L.IdUM = PR.IdUM AND PR.IdEmpresa = @PIdEmpresa'
            + @UnionesNivel
            + @UnionesAlmacen
            + N' WHERE L.IdEmpresa = @PIdEmpresa'
            + N' AND L.IdLista = @PIdLista'
            + N' AND (@PBusqueda = '''' OR P.GlsProducto LIKE @PBusquedaLike)'
            + N' AND (@PIdNivel = '''' OR P.IdNivel = @PIdNivel)'
            + N' ORDER BY L.IdProducto;';

        EXEC sys.sp_executesql
            @Sql,
            N'@PIdEmpresa CHAR(2), @PIdLista CHAR(8), @PIdNivel CHAR(8), @PBusqueda VARCHAR(255), @PBusquedaLike VARCHAR(257)',
            @PIdEmpresa = @IdEmpresa,
            @PIdLista = @IdLista,
            @PIdNivel = @IdNivel,
            @PBusqueda = @Busqueda,
            @PBusquedaLike = @BusquedaLike;
    END TRY
    BEGIN CATCH
        IF CURSOR_STATUS('local', 'CursorNiveles') >= -1
        BEGIN
            CLOSE CursorNiveles;
            DEALLOCATE CursorNiveles;
        END;

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
