
GO
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO

-- Firma: Codex - 10/06/2026 | Lista boletines deportivos registrados por el usuario autenticado para su perfil publico.
CREATE OR ALTER PROCEDURE dbo.Sp_BoletinesDeportivos_ListarPorUsuario
    @UsuarioId NVARCHAR(450)
AS
BEGIN
    SET NOCOUNT ON;

    BEGIN TRY
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
            b.FechaCreacion
        FROM dbo.BoletinesDeportivos b
        INNER JOIN dbo.UbigeoDistritos dist ON dist.CodigoUbigeo = b.CodigoUbigeo
        INNER JOIN dbo.UbigeoProvincias prov ON prov.CodigoProvincia = dist.CodigoProvincia
        INNER JOIN dbo.UbigeoDepartamentos dep ON dep.CodigoDepartamento = dist.CodigoDepartamento
        WHERE b.UsuarioId = @UsuarioId
        ORDER BY b.FechaCreacion DESC, b.IdBoletin DESC;
    END TRY
    BEGIN CATCH
        DECLARE @ErrorMessage NVARCHAR(4000), @ErrorSeverity INT, @ErrorState INT;
        SELECT @ErrorMessage = ERROR_MESSAGE(), @ErrorSeverity = ERROR_SEVERITY(), @ErrorState = ERROR_STATE();
        RAISERROR (@ErrorMessage, @ErrorSeverity, @ErrorState);
    END CATCH
END
GO
