-- Firma: FRANCO LARA | Create date: 17/11/2026 | Consulta las series de comprobantes desde Configuración HYT.
USE [SistemaWebHYT];
GO
CREATE OR ALTER PROCEDURE dbo.spu_SerieComprobanteHYT_Listar @UsuarioId NVARCHAR(128)
AS
BEGIN
    SET NOCOUNT ON;
    BEGIN TRY
        IF dbo.fn_HYTPermiso(@UsuarioId,N'CONFIGURACION.CONSULTAR')=0
            THROW 51600,N'[HYT_PERMISO] No tiene permiso para consultar la configuración.',1;
        SELECT COALESCE((SELECT serie.Tipo,serie.Serie FROM dbo.SerieComprobanteHYT AS serie
            ORDER BY serie.Tipo,serie.Serie FOR JSON PATH),N'[]') AS Datos;
    END TRY
    BEGIN CATCH
        DECLARE @ErrorMessage NVARCHAR(4000),@ErrorSeverity INT,@ErrorState INT;
        SELECT @ErrorMessage=ERROR_MESSAGE(),@ErrorSeverity=ERROR_SEVERITY(),@ErrorState=ERROR_STATE();
        RAISERROR (@ErrorMessage,@ErrorSeverity,@ErrorState);
    END CATCH
END;
GO
GRANT EXECUTE ON OBJECT::dbo.spu_SerieComprobanteHYT_Listar TO hyt_app_executor;
GO
