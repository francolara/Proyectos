-- =============================================
-- Author:      FRANCO LARA / Codex
-- Create date: 05/11/2026
-- Description: Acepta % y * y busca todos los terminos en cualquier orden.
-- =============================================
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
        DECLARE @Termino VARCHAR(500);
        DECLARE @PosSeparador INT;
        DECLARE @FiltroBusqueda NVARCHAR(MAX) = N'';

        SET @Busqueda = REPLACE(REPLACE(REPLACE(LTRIM(RTRIM(ISNULL(@Busqueda, ''))), '*', ' '), '%', ' '), CHAR(9), ' ');

        WHILE LEN(@Busqueda) > 0
        BEGIN
            SET @PosSeparador = CHARINDEX(' ', @Busqueda + ' ');
            SET @Termino = LEFT(@Busqueda, @PosSeparador - 1);
            SET @Busqueda = LTRIM(SUBSTRING(@Busqueda, @PosSeparador + 1, 250));

            IF @Termino <> ''
            BEGIN
                -- Escapar comillas antes de incorporar el termino al SQL dinamico.
                SET @Termino = REPLACE(@Termino, '''', '''''');
                SET @FiltroBusqueda = @FiltroBusqueda
                    + N'AND (CHARINDEX(''' + @Termino + N''', P.GlsProducto) > 0 '
                    + N'OR CHARINDEX(''' + @Termino + N''', P.CodigoRapido) > 0 '
                    + N'OR CHARINDEX(''' + @Termino + N''', P.IdFabricante) > 0 '
                    + N'OR CHARINDEX(''' + @Termino + N''', P.IdProducto) > 0) ';
            END;
        END;

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

        SET @Sql = @Sql + @FiltroBusqueda;

        IF @OrdenarPorCodigoRapido = 1
            SET @Sql = @Sql + N'ORDER BY P.CodigoRapido;';
        ELSE
            SET @Sql = @Sql + N'ORDER BY P.IdProducto;';

        EXEC sys.sp_executesql
            @Sql,
            N'@PIdEmpresa CHAR(2)',
            @PIdEmpresa = @IdEmpresa;
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
