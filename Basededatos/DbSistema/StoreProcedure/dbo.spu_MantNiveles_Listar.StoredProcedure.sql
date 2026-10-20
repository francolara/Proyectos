-- =============================================
-- Author:      FRANCO LARA / Codex
-- Create date: 15/10/2026
-- Description: Lista niveles y sus predecesores para el mantenimiento.
-- =============================================
-- =============================================
-- Author:      FRANCO LARA / Codex
-- Create date: 16/10/2026
-- Description: Permite listar el primer nivel cuando no existen predecesores.
-- =============================================
CREATE OR ALTER PROCEDURE dbo.spu_MantNiveles_Listar
    @IdEmpresa CHAR(2),
    @IdTipoNivel CHAR(8),
    @Busqueda VARCHAR(255) = ''
AS
BEGIN
    SET NOCOUNT ON;

    BEGIN TRY
        DECLARE @PesoActual INT, @PesoPadre INT, @Indice INT = 0;
        DECLARE @Campos NVARCHAR(MAX) = N'', @Uniones NVARCHAR(MAX) = N'';
        DECLARE @Sql NVARCHAR(MAX);

        SET @Busqueda = LTRIM(RTRIM(ISNULL(@Busqueda, '')));

        SELECT @PesoActual = TN.Peso
        FROM dbo.TiposNiveles AS TN
        WHERE TN.IdEmpresa = @IdEmpresa
          AND TN.IdTipoNivel = @IdTipoNivel;

        IF @PesoActual IS NULL
            RETURN;

        SET @PesoPadre = @PesoActual - 1;

        WHILE @PesoPadre >= 1 AND @Indice < 4
        BEGIN
            SET @Indice += 1;
            SET @Campos += N', N' + RIGHT(N'00' + CONVERT(NVARCHAR(2), @Indice), 2)
                        + N'.IdNivel AS IdNivel' + RIGHT(N'00' + CONVERT(NVARCHAR(2), @Indice), 2)
                        + N', N' + RIGHT(N'00' + CONVERT(NVARCHAR(2), @Indice), 2)
                        + N'.GlsNivel AS GlsNivel' + RIGHT(N'00' + CONVERT(NVARCHAR(2), @Indice), 2);

            IF @Indice = 1
                SET @Uniones += N' INNER JOIN dbo.Niveles AS N01 ON N.IdNivelPred = N01.IdNivel AND N01.IdEmpresa = N.IdEmpresa';
            ELSE
                SET @Uniones += N' INNER JOIN dbo.Niveles AS N' + RIGHT(N'00' + CONVERT(NVARCHAR(2), @Indice), 2)
                            + N' ON N' + RIGHT(N'00' + CONVERT(NVARCHAR(2), @Indice - 1), 2)
                            + N'.IdNivelPred = N' + RIGHT(N'00' + CONVERT(NVARCHAR(2), @Indice), 2)
                            + N'.IdNivel AND N' + RIGHT(N'00' + CONVERT(NVARCHAR(2), @Indice), 2) + N'.IdEmpresa = N.IdEmpresa';

            SET @PesoPadre -= 1;
        END;

        SET @Sql = N'SELECT '
                 + CASE WHEN @Campos = N'' THEN N'' ELSE STUFF(@Campos, 1, 2, N'') + N',' END
                 + N' N.IdNivel, N.GlsNivel
                    FROM dbo.Niveles AS N'
                 + @Uniones
                 + N' WHERE N.IdEmpresa = @PEmpresa
                       AND N.IdTipoNivel = @PTipo'
                 + CASE WHEN @Busqueda = '' THEN N'' ELSE N' AND N.GlsNivel LIKE ''%'' + @PBusqueda + ''%''' END
                 + N' ORDER BY N.IdNivel;';

        EXEC sys.sp_executesql
            @Sql,
            N'@PEmpresa CHAR(2), @PTipo CHAR(8), @PBusqueda VARCHAR(255)',
            @PEmpresa = @IdEmpresa,
            @PTipo = @IdTipoNivel,
            @PBusqueda = @Busqueda;
    END TRY
    BEGIN CATCH
        DECLARE @ErrorMessage NVARCHAR(4000), @ErrorSeverity INT, @ErrorState INT;

        SELECT @ErrorMessage = ERROR_MESSAGE(), @ErrorSeverity = ERROR_SEVERITY(), @ErrorState = ERROR_STATE();
        RAISERROR (@ErrorMessage, @ErrorSeverity, @ErrorState);
    END CATCH
END
