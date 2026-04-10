USE [DbSportCenter]
GO
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
-- =============================================
-- Author:        FRANCO LARA
-- Create date:   09/04/2026
-- Description:   Combo de tipos de documento de comprobante habilitados por negocio.
-- =============================================
CREATE OR ALTER PROCEDURE dbo.Sp_Combos_DocumentosComprobanteNegocio
    @NegocioId INT,
    @Tributario BIT = NULL
AS
BEGIN
    SET NOCOUNT ON;
    BEGIN TRY
        SELECT
            t.CodigoSunat AS Value,
            CONCAT(t.Nombre, N' (', t.CodigoSunat, N')') AS Text
        FROM dbo.NegociosTiposDocumentoComprobante ntd
        INNER JOIN dbo.TiposDocumentoComprobanteSuperMaestro t ON t.CodigoSunat = ntd.CodigoSunat
        WHERE ntd.NegocioId = @NegocioId
          AND ntd.Activo = 1
          AND t.Activo = 1
          AND t.Habilitado = 1
          AND (@Tributario IS NULL OR t.Tributario = @Tributario)
        ORDER BY t.Orden, t.CodigoSunat;
    END TRY
    BEGIN CATCH
        DECLARE @ErrorMessage NVARCHAR(4000), @ErrorSeverity INT, @ErrorState INT;
        SELECT @ErrorMessage = ERROR_MESSAGE(), @ErrorSeverity = ERROR_SEVERITY(), @ErrorState = ERROR_STATE();
        RAISERROR (@ErrorMessage, @ErrorSeverity, @ErrorState);
    END CATCH
END
GO
