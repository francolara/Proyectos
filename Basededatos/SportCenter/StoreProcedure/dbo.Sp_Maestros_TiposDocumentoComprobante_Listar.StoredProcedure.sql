USE [DbSportCenter]
GO
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
-- =============================================
-- Author:        FRANCO LARA
-- Create date:   09/04/2026
-- Description:   Lista tipos de documento de comprobante configurados para el negocio.
-- =============================================
CREATE OR ALTER PROCEDURE dbo.Sp_Maestros_TiposDocumentoComprobante_Listar
    @NegocioId INT
AS
BEGIN
    SET NOCOUNT ON;
    BEGIN TRY
        SELECT
            ntd.Id,
            ntd.CodigoSunat,
            t.Nombre,
            t.Tributario,
            t.Habilitado,
            ntd.Activo
        FROM dbo.NegociosTiposDocumentoComprobante ntd
        INNER JOIN dbo.TiposDocumentoComprobanteSuperMaestro t ON t.CodigoSunat = ntd.CodigoSunat
        WHERE ntd.NegocioId = @NegocioId
        ORDER BY t.Orden, t.CodigoSunat;
    END TRY
    BEGIN CATCH
        DECLARE @ErrorMessage NVARCHAR(4000), @ErrorSeverity INT, @ErrorState INT;
        SELECT @ErrorMessage = ERROR_MESSAGE(), @ErrorSeverity = ERROR_SEVERITY(), @ErrorState = ERROR_STATE();
        RAISERROR (@ErrorMessage, @ErrorSeverity, @ErrorState);
    END CATCH
END
GO
