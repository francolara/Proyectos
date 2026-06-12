
GO
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO

-- Firma: Codex - 10/06/2026 | Lista boletines deportivos registrados por el usuario autenticado para su perfil publico.
-- Firma: Codex - 11/06/2026 | Agrega paginacion server-side de 5 registros para Mis boletines devolviendo total de filas por usuario.
CREATE OR ALTER PROCEDURE dbo.Sp_BoletinesDeportivos_ListarPorUsuario
    @UsuarioId NVARCHAR(450),
    @Pagina INT = 1,
    @TamanoPagina INT = 5
AS
BEGIN
    SET NOCOUNT ON;

    BEGIN TRY
        DECLARE @PaginaNorm INT = CASE WHEN ISNULL(@Pagina, 0) < 1 THEN 1 ELSE @Pagina END;
        DECLARE @TamanoPaginaNorm INT = CASE WHEN ISNULL(@TamanoPagina, 0) < 1 THEN 5 ELSE @TamanoPagina END;
        DECLARE @Offset INT = (@PaginaNorm - 1) * @TamanoPaginaNorm;

        ;WITH BoletinesFiltrados AS
        (
            SELECT
                b.IdBoletin,
                b.Titulo,
                b.Descripcion,
                b.ImagenUrl,
                b.FechaEvento,
                b.CodigoUbigeo,
                dep.Nombre AS Departamento,
                prov.Nombre AS Provincia,
                dist.Nombre AS Distrito,
                dist.Zona,
                b.TipoRegistro,
                b.Activo,
                b.FechaCreacion,
                COUNT(1) OVER() AS TotalRegistros
            FROM dbo.BoletinesDeportivos b
            INNER JOIN dbo.UbigeoDistritos dist ON dist.CodigoUbigeo = b.CodigoUbigeo
            INNER JOIN dbo.UbigeoProvincias prov ON prov.CodigoProvincia = dist.CodigoProvincia
            INNER JOIN dbo.UbigeoDepartamentos dep ON dep.CodigoDepartamento = dist.CodigoDepartamento
            WHERE b.UsuarioId = @UsuarioId
        )
        SELECT
            IdBoletin,
            Titulo,
            Descripcion,
            ImagenUrl,
            FechaEvento,
            CodigoUbigeo,
            Departamento,
            Provincia,
            Distrito,
            Zona,
            TipoRegistro,
            Activo,
            FechaCreacion,
            TotalRegistros
        FROM BoletinesFiltrados
        ORDER BY FechaCreacion DESC, IdBoletin DESC
        OFFSET @Offset ROWS FETCH NEXT @TamanoPaginaNorm ROWS ONLY;
    END TRY
    BEGIN CATCH
        DECLARE @ErrorMessage NVARCHAR(4000), @ErrorSeverity INT, @ErrorState INT;
        SELECT @ErrorMessage = ERROR_MESSAGE(), @ErrorSeverity = ERROR_SEVERITY(), @ErrorState = ERROR_STATE();
        RAISERROR (@ErrorMessage, @ErrorSeverity, @ErrorState);
    END CATCH
END
GO
