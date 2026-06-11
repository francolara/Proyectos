
GO
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO

-- Firma: Codex - 10/06/2026 | Obtiene el detalle ampliado de un boletin deportivo para vista completa y administracion.
CREATE OR ALTER PROCEDURE dbo.Sp_BoletinesDeportivos_ObtenerPorId
    @IdBoletin INT
AS
BEGIN
    SET NOCOUNT ON;

    BEGIN TRY
        SELECT
            b.IdBoletin,
            b.UsuarioId,
            b.PerfilPublicoId,
            b.Titulo,
            b.Descripcion,
            b.ImagenUrl,
            b.FechaEvento,
            b.CodigoUbigeo,
            dep.CodigoDepartamento,
            prov.CodigoProvincia,
            dep.Nombre AS Departamento,
            prov.Nombre AS Provincia,
            dist.Nombre AS Distrito,
            dist.Zona,
            b.TipoRegistro,
            b.Activo,
            b.FechaCreacion,
            b.UsuarioCreacion,
            b.FechaActualizacion,
            b.UsuarioActualizacion
        FROM dbo.BoletinesDeportivos b
        INNER JOIN dbo.UbigeoDistritos dist ON dist.CodigoUbigeo = b.CodigoUbigeo
        INNER JOIN dbo.UbigeoProvincias prov ON prov.CodigoProvincia = dist.CodigoProvincia
        INNER JOIN dbo.UbigeoDepartamentos dep ON dep.CodigoDepartamento = dist.CodigoDepartamento
        WHERE b.IdBoletin = @IdBoletin;
    END TRY
    BEGIN CATCH
        DECLARE @ErrorMessage NVARCHAR(4000), @ErrorSeverity INT, @ErrorState INT;
        SELECT @ErrorMessage = ERROR_MESSAGE(), @ErrorSeverity = ERROR_SEVERITY(), @ErrorState = ERROR_STATE();
        RAISERROR (@ErrorMessage, @ErrorSeverity, @ErrorState);
    END CATCH
END
GO
