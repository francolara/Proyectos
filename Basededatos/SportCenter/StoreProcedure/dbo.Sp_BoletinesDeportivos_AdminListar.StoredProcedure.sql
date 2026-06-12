
GO
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO

-- Firma: Codex - 11/06/2026 | Agrega paginacion server-side para el listado de boletines del super admin usando pagina, tamano y total de registros.
CREATE OR ALTER PROCEDURE dbo.Sp_BoletinesDeportivos_AdminListar
    @Activo BIT = NULL,
    @TipoRegistro CHAR(1) = NULL,
    @CodigoDepartamento CHAR(2) = NULL,
    @CodigoProvincia CHAR(4) = NULL,
    @CodigoUbigeo CHAR(6) = NULL,
    @Zona NVARCHAR(20) = NULL,
    @Anio INT = NULL,
    @Mes INT = NULL,
    @Pagina INT = 1,
    @TamanoPagina INT = 5
AS
BEGIN
    SET NOCOUNT ON;

    BEGIN TRY
        DECLARE @ZonaNorm NVARCHAR(20) = NULLIF(LTRIM(RTRIM(@Zona)), N'');
        DECLARE @TipoRegistroNorm CHAR(1) = UPPER(ISNULL(NULLIF(LTRIM(RTRIM(@TipoRegistro)), ''), ''));
        DECLARE @PaginaNorm INT = CASE WHEN ISNULL(@Pagina, 0) < 1 THEN 1 ELSE @Pagina END;
        DECLARE @TamanoPaginaNorm INT = CASE WHEN ISNULL(@TamanoPagina, 0) < 1 THEN 5 ELSE @TamanoPagina END;
        DECLARE @Offset INT = (@PaginaNorm - 1) * @TamanoPaginaNorm;

        ;WITH BoletinesFiltrados AS
        (
            SELECT
                b.IdBoletin,
                b.UsuarioId,
                au.Email,
                COALESCE(NULLIF(LTRIM(RTRIM(CONCAT(up.Nombres, N' ', up.Apellidos))), N''), au.UserName, au.Email, b.UsuarioId) AS NombreAutor,
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
            INNER JOIN dbo.AspNetUsers au ON au.Id = b.UsuarioId
            LEFT JOIN dbo.UsuariosPublicosPerfil up ON up.UsuarioId = b.UsuarioId
            INNER JOIN dbo.UbigeoDistritos dist ON dist.CodigoUbigeo = b.CodigoUbigeo
            INNER JOIN dbo.UbigeoProvincias prov ON prov.CodigoProvincia = dist.CodigoProvincia
            INNER JOIN dbo.UbigeoDepartamentos dep ON dep.CodigoDepartamento = dist.CodigoDepartamento
            WHERE (@Activo IS NULL OR b.Activo = @Activo)
              AND (@TipoRegistroNorm = '' OR b.TipoRegistro = @TipoRegistroNorm)
              AND (@CodigoDepartamento IS NULL OR dist.CodigoDepartamento = @CodigoDepartamento)
              AND (@CodigoProvincia IS NULL OR dist.CodigoProvincia = @CodigoProvincia)
              AND (@CodigoUbigeo IS NULL OR dist.CodigoUbigeo = @CodigoUbigeo)
              AND (@ZonaNorm IS NULL OR dist.Zona = @ZonaNorm)
              AND (@Anio IS NULL OR YEAR(b.FechaEvento) = @Anio)
              AND (@Mes IS NULL OR MONTH(b.FechaEvento) = @Mes)
        )
        SELECT
            IdBoletin,
            UsuarioId,
            Email,
            NombreAutor,
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
        ORDER BY FechaEvento DESC, FechaCreacion DESC, IdBoletin DESC
        OFFSET @Offset ROWS FETCH NEXT @TamanoPaginaNorm ROWS ONLY;
    END TRY
    BEGIN CATCH
        DECLARE @ErrorMessage NVARCHAR(4000), @ErrorSeverity INT, @ErrorState INT;
        SELECT @ErrorMessage = ERROR_MESSAGE(), @ErrorSeverity = ERROR_SEVERITY(), @ErrorState = ERROR_STATE();
        RAISERROR (@ErrorMessage, @ErrorSeverity, @ErrorState);
    END CATCH
END
GO
