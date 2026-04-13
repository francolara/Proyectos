/*
Firma: Codex - 10/04/2026
Descripcion: Crea/actualiza SP para obtener ValorParametro por NombreParametro en ParametrosGlobales.
*/
USE [DbSportCenter]
GO

CREATE OR ALTER PROCEDURE dbo.Sp_ParametrosGlobales_ObtenerValor
    @NombreParametro NVARCHAR(100)
AS
BEGIN
    SET NOCOUNT ON;

    BEGIN TRY
        SELECT TOP (1)
            p.ValorParametro
        FROM dbo.ParametrosGlobales p
        WHERE p.NombreParametro = @NombreParametro;
    END TRY
    BEGIN CATCH
        DECLARE @ErrorMessage NVARCHAR(4000), @ErrorSeverity INT, @ErrorState INT;
        SELECT @ErrorMessage = ERROR_MESSAGE(), @ErrorSeverity = ERROR_SEVERITY(), @ErrorState = ERROR_STATE();
        RAISERROR (@ErrorMessage, @ErrorSeverity, @ErrorState);
    END CATCH
END
GO
