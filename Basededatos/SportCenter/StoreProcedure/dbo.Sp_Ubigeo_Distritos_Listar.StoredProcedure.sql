
GO
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO

-- Firma: Codex - 04/04/2026 | Actualizacion individual de Sp_Ubigeo_Distritos_Listar por integracion de ubigeo fiscal.
-- Firma: Codex - 11/06/2026 | Agrega filtro opcional por zona para reutilizar distritos en formularios publicos y desafios sin romper llamadas existentes.
CREATE OR ALTER PROCEDURE dbo.Sp_Ubigeo_Distritos_Listar
    @CodigoProvincia CHAR(4),
    @Zona NVARCHAR(30) = NULL
AS
BEGIN
    SET NOCOUNT ON;
    BEGIN TRY
        SELECT
            d.CodigoUbigeo,
            d.Nombre
        FROM dbo.UbigeoDistritos d
        WHERE d.Activo = 1
          AND d.CodigoProvincia = @CodigoProvincia
          AND (
                @Zona IS NULL
                OR LTRIM(RTRIM(@Zona)) = N''
                OR ISNULL(d.Zona, N'') = LTRIM(RTRIM(@Zona))
              )
        ORDER BY d.Nombre;
    END TRY
    BEGIN CATCH
        DECLARE @ErrorMessage NVARCHAR(4000), @ErrorSeverity INT, @ErrorState INT;
        SELECT @ErrorMessage = ERROR_MESSAGE(), @ErrorSeverity = ERROR_SEVERITY(), @ErrorState = ERROR_STATE();
        RAISERROR (@ErrorMessage, @ErrorSeverity, @ErrorState);
    END CATCH
END
GO
