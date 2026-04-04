USE [DbSportCenter]
GO
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO

-- Firma: Codex - 04/04/2026 | Actualizacion individual de Sp_Ubigeo_Distritos_Listar por integracion de ubigeo fiscal.
CREATE OR ALTER PROCEDURE dbo.Sp_Ubigeo_Distritos_Listar
    @CodigoProvincia CHAR(4)
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
        ORDER BY d.Nombre;
    END TRY
    BEGIN CATCH
        DECLARE @ErrorMessage NVARCHAR(4000), @ErrorSeverity INT, @ErrorState INT;
        SELECT @ErrorMessage = ERROR_MESSAGE(), @ErrorSeverity = ERROR_SEVERITY(), @ErrorState = ERROR_STATE();
        RAISERROR (@ErrorMessage, @ErrorSeverity, @ErrorState);
    END CATCH
END
GO
