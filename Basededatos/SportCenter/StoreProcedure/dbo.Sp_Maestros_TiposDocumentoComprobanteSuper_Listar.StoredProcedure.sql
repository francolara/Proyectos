USE [DbSportCenter]
GO
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
-- =============================================
-- Author:        FRANCO LARA
-- Create date:   09/04/2026
-- Description:   Lista tipos de documento del supermaestro para mantenimiento por negocio.
-- =============================================
CREATE OR ALTER PROCEDURE dbo.Sp_Maestros_TiposDocumentoComprobanteSuper_Listar
AS
BEGIN
    SET NOCOUNT ON;
    BEGIN TRY
        SELECT
            t.CodigoSunat AS Value,
            CONCAT(t.Nombre, N' (', t.CodigoSunat, N')') AS Text
        FROM dbo.TiposDocumentoComprobanteSuperMaestro t
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
