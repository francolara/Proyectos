
GO
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO

-- Firma: Codex - 10/06/2026 | Lista zonas geograficas registradas en ubigeo distritos para filtros publicos de boletines.
CREATE OR ALTER PROCEDURE dbo.Sp_Ubigeo_Zonas_Listar
    @CodigoDepartamento CHAR(2) = NULL,
    @CodigoProvincia CHAR(4) = NULL
AS
BEGIN
    SET NOCOUNT ON;

    BEGIN TRY
        SELECT DISTINCT
            d.Zona AS Value,
            d.Zona AS Text
        FROM dbo.UbigeoDistritos d
        WHERE d.Activo = 1
          AND NULLIF(LTRIM(RTRIM(d.Zona)), N'') IS NOT NULL
          AND (@CodigoDepartamento IS NULL OR d.CodigoDepartamento = @CodigoDepartamento)
          AND (@CodigoProvincia IS NULL OR d.CodigoProvincia = @CodigoProvincia)
        ORDER BY d.Zona;
    END TRY
    BEGIN CATCH
        DECLARE @ErrorMessage NVARCHAR(4000), @ErrorSeverity INT, @ErrorState INT;
        SELECT @ErrorMessage = ERROR_MESSAGE(), @ErrorSeverity = ERROR_SEVERITY(), @ErrorState = ERROR_STATE();
        RAISERROR (@ErrorMessage, @ErrorSeverity, @ErrorState);
    END CATCH
END
GO
