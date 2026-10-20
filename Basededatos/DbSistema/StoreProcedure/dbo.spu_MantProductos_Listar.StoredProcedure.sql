-- =============================================
-- Author:      FRANCO LARA / Codex
-- Create date: 15/10/2026
-- Description: Lista productos con sus niveles dinamicos para el mantenimiento de productos.
-- =============================================
CREATE OR ALTER PROCEDURE dbo.spu_MantProductos_Listar
    @IdEmpresa CHAR(2),
    @Busqueda VARCHAR(250) = '',
    @OrdenarPorCodigoRapido BIT = 0
AS
BEGIN
    SET NOCOUNT ON;

    BEGIN TRY
        DECLARE @Sql NVARCHAR(MAX);
        DECLARE @ColumnasNivel NVARCHAR(MAX) = N'';
        DECLARE @UnionesNivel NVARCHAR(MAX) = N'';
        DECLARE @AliasActual SYSNAME;
        DECLARE @AliasAnterior SYSNAME = N'';
        DECLARE @Indice INT = 0;
        DECLARE @Peso INT;
        DECLARE @BusquedaLike VARCHAR(252);

        SET @Busqueda = LTRIM(RTRIM(ISNULL(@Busqueda, '')));
        SET @BusquedaLike = '%' + @Busqueda + '%';

        DECLARE CursorNiveles CURSOR LOCAL FAST_FORWARD FOR
            SELECT TN.Peso
            FROM dbo.TiposNiveles AS TN
            WHERE TN.IdEmpresa = @IdEmpresa
            ORDER BY TN.Peso DESC;

        OPEN CursorNiveles;
        FETCH NEXT FROM CursorNiveles INTO @Peso;

        WHILE @@FETCH_STATUS = 0
        BEGIN
            SET @Indice = @Indice + 1;
            SET @AliasActual = N'N' + RIGHT(N'00' + CONVERT(NVARCHAR(10), @Indice), 2);

            SET @ColumnasNivel = @ColumnasNivel
                + @AliasActual + N'.IdNivel AS IdNivel' + RIGHT(N'00' + CONVERT(NVARCHAR(10), @Indice), 2) + N','
                + @AliasActual + N'.GlsNivel AS GlsNivel' + RIGHT(N'00' + CONVERT(NVARCHAR(10), @Indice), 2) + N',';

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

        SET @Sql = N'SELECT ' + @ColumnasNivel
            + N'P.IdProducto, P.GlsProducto, M.GlsMarca, UMC.AbreUM AS GlsUMC, UMV.AbreUM AS GlsUMV, '
            + N'CASE WHEN P.EstProducto = ''A'' THEN ''ACT'' ELSE ''INA'' END AS EstProducto, '
            + N'P.IdFabricante, P.CtaContable2, P.CtaContable, P.CtaContable_Relacionada, P.CodigoRapido '
            + N'FROM dbo.Productos AS P '
            + N'LEFT JOIN dbo.Marcas AS M ON P.IdEmpresa = M.IdEmpresa AND P.IdMarca = M.IdMarca '
            + N'LEFT JOIN dbo.UnidadMedida AS UMC ON P.IdUMCompra = UMC.IdUM '
            + N'LEFT JOIN dbo.UnidadMedida AS UMV ON P.IdUMVenta = UMV.IdUM '
            + @UnionesNivel
            + N' WHERE P.IdEmpresa = @PIdEmpresa ';

        IF @Busqueda <> ''
            SET @Sql = @Sql + N'AND (P.GlsProducto LIKE @PBusquedaLike '
                + N'OR P.CodigoRapido LIKE @PBusquedaLike '
                + N'OR P.IdFabricante LIKE @PBusquedaLike '
                + N'OR P.IdProducto LIKE @PBusquedaLike) ';

        IF @OrdenarPorCodigoRapido = 1
            SET @Sql = @Sql + N'ORDER BY P.CodigoRapido;';
        ELSE
            SET @Sql = @Sql + N'ORDER BY P.IdProducto;';

        EXEC sys.sp_executesql
            @Sql,
            N'@PIdEmpresa CHAR(2), @PBusquedaLike VARCHAR(252)',
            @PIdEmpresa = @IdEmpresa,
            @PBusquedaLike = @BusquedaLike;
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
