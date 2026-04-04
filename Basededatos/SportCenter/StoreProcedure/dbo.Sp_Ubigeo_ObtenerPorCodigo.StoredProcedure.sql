USE [DbSportCenter]
GO
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO

-- Firma: Codex - 04/04/2026 | Actualizacion individual de Sp_Ubigeo_ObtenerPorCodigo por integracion de ubigeo fiscal.
CREATE OR ALTER PROCEDURE dbo.Sp_Ubigeo_ObtenerPorCodigo
    @CodigoUbigeo CHAR(6)
AS
BEGIN
    SET NOCOUNT ON;
    BEGIN TRY
        SELECT
            d.CodigoUbigeo,
            d.CodigoDepartamento,
            d.CodigoProvincia,
            dep.Nombre AS Departamento,
            prov.Nombre AS Provincia,
            d.Nombre AS Distrito
        FROM dbo.UbigeoDistritos d
        INNER JOIN dbo.UbigeoDepartamentos dep ON dep.CodigoDepartamento = d.CodigoDepartamento
        INNER JOIN dbo.UbigeoProvincias prov ON prov.CodigoProvincia = d.CodigoProvincia
        WHERE d.CodigoUbigeo = @CodigoUbigeo
          AND d.Activo = 1
          AND dep.Activo = 1
          AND prov.Activo = 1;
    END TRY
    BEGIN CATCH
        DECLARE @ErrorMessage NVARCHAR(4000), @ErrorSeverity INT, @ErrorState INT;
        SELECT @ErrorMessage = ERROR_MESSAGE(), @ErrorSeverity = ERROR_SEVERITY(), @ErrorState = ERROR_STATE();
        RAISERROR (@ErrorMessage, @ErrorSeverity, @ErrorState);
    END CATCH
END
GO
