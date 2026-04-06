USE [DbSportCenter]
GO
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO

-- Firma: Codex - 04/04/2026 | Combo centralizado de tipos de documento SUNAT para clientes y configuracion.
CREATE OR ALTER PROCEDURE dbo.Sp_Combos_TiposDocumentoIdentidadSunat
AS
BEGIN
    SET NOCOUNT ON;

    BEGIN TRY
        SELECT
            t.CodigoSunat,
            CONCAT(t.CodigoInterno, N' - ', t.Nombre, N' (', t.CodigoSunat, N')') AS Nombre
        FROM dbo.TiposDocumentoIdentidadSunat t
        WHERE t.Activo = 1
        ORDER BY t.Orden, t.CodigoSunat;
    END TRY
    BEGIN CATCH
        DECLARE @ErrorMessage NVARCHAR(4000), @ErrorSeverity INT, @ErrorState INT;
        SELECT @ErrorMessage = ERROR_MESSAGE(), @ErrorSeverity = ERROR_SEVERITY(), @ErrorState = ERROR_STATE();
        RAISERROR (@ErrorMessage, @ErrorSeverity, @ErrorState);
    END CATCH
END
GO
