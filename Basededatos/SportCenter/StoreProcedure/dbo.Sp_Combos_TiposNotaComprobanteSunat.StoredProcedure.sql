USE [DbSportCenter]
GO
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
-- =============================================
-- Author:        FRANCO LARA
-- Create date:   11/04/2026
-- Description:   Combo de tipos SUNAT para nota de credito/debito (07/08).
-- =============================================
CREATE OR ALTER PROCEDURE dbo.Sp_Combos_TiposNotaComprobanteSunat
    @TipoNota CHAR(2)
AS
BEGIN
    SET NOCOUNT ON;
    BEGIN TRY
        DECLARE @TipoNotaNorm CHAR(2) = UPPER(LTRIM(RTRIM(ISNULL(@TipoNota, ''))));

        IF @TipoNotaNorm = 'NC' SET @TipoNotaNorm = '07';
        IF @TipoNotaNorm = 'ND' SET @TipoNotaNorm = '08';

        IF @TipoNotaNorm NOT IN ('07', '08')
            RAISERROR('Tipo de nota no valido.', 16, 1);

        SELECT
            t.CodigoSunat AS Value,
            CONCAT(t.CodigoSunat, N' - ', t.Nombre) AS Text
        FROM dbo.TiposNotaComprobanteSunat t
        WHERE t.TipoNota = @TipoNotaNorm
          AND t.Activo = 1
        ORDER BY t.Orden, t.CodigoSunat;
    END TRY
    BEGIN CATCH
        DECLARE @ErrorMessage NVARCHAR(4000), @ErrorSeverity INT, @ErrorState INT;
        SELECT @ErrorMessage = ERROR_MESSAGE(), @ErrorSeverity = ERROR_SEVERITY(), @ErrorState = ERROR_STATE();
        RAISERROR (@ErrorMessage, @ErrorSeverity, @ErrorState);
    END CATCH
END
GO
