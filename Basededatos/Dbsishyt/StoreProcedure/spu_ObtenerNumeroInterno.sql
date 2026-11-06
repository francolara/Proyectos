-- =============================================
-- Author:        FRANCO LARA
-- Create date:   03/11/2026
-- Description:   Reserva numero interno dentro de la transaccion del registro.
-- =============================================
USE [SistemaWebHYT];
GO
CREATE OR ALTER PROCEDURE dbo.spu_ObtenerNumeroInterno
    @TipoRegistro NVARCHAR(40),
    @NumeroInterno NVARCHAR(40) OUTPUT
AS
BEGIN
    SET NOCOUNT ON;
    BEGIN TRY
        IF @@TRANCOUNT = 0 OR XACT_STATE() <> 1
            THROW 51011, N'La numeracion requiere la transaccion del registro.', 1;
        DECLARE @Prefijo NVARCHAR(20), @Ultimo BIGINT;
        SELECT @Prefijo = numeracion.Prefijo, @Ultimo = numeracion.UltimoNumero
        FROM dbo.NumeracionInterna AS numeracion WITH (UPDLOCK, HOLDLOCK)
        WHERE numeracion.TipoRegistro = @TipoRegistro;
        IF @Ultimo IS NULL OR @Ultimo = 9223372036854775807
            THROW 51012, N'Numeracion inexistente o agotada.', 1;
        UPDATE dbo.NumeracionInterna SET UltimoNumero = @Ultimo + 1 WHERE TipoRegistro = @TipoRegistro;
        SET @NumeroInterno = CONCAT(@Prefijo, CONVERT(NVARCHAR(20), @Ultimo + 1));
    END TRY
    BEGIN CATCH
        DECLARE @ErrorMessage NVARCHAR(4000), @ErrorSeverity INT, @ErrorState INT;
        SELECT @ErrorMessage = ERROR_MESSAGE(), @ErrorSeverity = ERROR_SEVERITY(), @ErrorState = ERROR_STATE();
        RAISERROR (@ErrorMessage, @ErrorSeverity, @ErrorState);
    END CATCH
END;
GO
