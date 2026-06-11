
GO
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO

-- Firma: Codex - 10/06/2026 | Lista boletines deportivos visibles en el home con filtros por ubigeo, zona y periodo.
CREATE OR ALTER PROCEDURE dbo.Sp_BoletinesDeportivos_ListarPublico
    @CodigoDepartamento CHAR(2) = NULL,
    @CodigoProvincia CHAR(4) = NULL,
    @CodigoUbigeo CHAR(6) = NULL,
    @Zona NVARCHAR(20) = NULL,
    @Anio INT = NULL,
    @Mes INT = NULL
AS
BEGIN
    SET NOCOUNT ON;

    BEGIN TRY
        DECLARE @ZonaNorm NVARCHAR(20) = NULLIF(LTRIM(RTRIM(@Zona)), N'');

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
            b.FechaCreacion
        FROM dbo.BoletinesDeportivos b
        INNER JOIN dbo.UbigeoDistritos dist ON dist.CodigoUbigeo = b.CodigoUbigeo AND dist.Activo = 1
        INNER JOIN dbo.UbigeoProvincias prov ON prov.CodigoProvincia = dist.CodigoProvincia AND prov.Activo = 1
        INNER JOIN dbo.UbigeoDepartamentos dep ON dep.CodigoDepartamento = dist.CodigoDepartamento AND dep.Activo = 1
        WHERE b.Activo = 1
          AND (@CodigoDepartamento IS NULL OR dist.CodigoDepartamento = @CodigoDepartamento)
          AND (@CodigoProvincia IS NULL OR dist.CodigoProvincia = @CodigoProvincia)
          AND (@CodigoUbigeo IS NULL OR dist.CodigoUbigeo = @CodigoUbigeo)
          AND (@ZonaNorm IS NULL OR dist.Zona = @ZonaNorm)
          AND (@Anio IS NULL OR YEAR(b.FechaEvento) = @Anio)
          AND (@Mes IS NULL OR MONTH(b.FechaEvento) = @Mes)
        ORDER BY b.FechaEvento DESC, b.FechaCreacion DESC, b.IdBoletin DESC;
    END TRY
    BEGIN CATCH
        DECLARE @ErrorMessage NVARCHAR(4000), @ErrorSeverity INT, @ErrorState INT;
        SELECT @ErrorMessage = ERROR_MESSAGE(), @ErrorSeverity = ERROR_SEVERITY(), @ErrorState = ERROR_STATE();
        RAISERROR (@ErrorMessage, @ErrorSeverity, @ErrorState);
    END CATCH
END
GO
