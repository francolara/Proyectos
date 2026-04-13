USE [DbSportCenter]
GO
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO

-- Firma: Codex - 10/04/2026 | Retorna valor de parametro global por NombreParametro.
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
