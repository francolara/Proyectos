-- Firma: FRANCO LARA | Create date: 17/11/2026 | Consulta nombres, imágenes y firma predeterminada.
USE [SistemaWebHYT];
GO
CREATE OR ALTER PROCEDURE dbo.spu_FirmaPresupuestoHYT_Listar @UsuarioId NVARCHAR(128)
AS
BEGIN
    SET NOCOUNT ON;
    BEGIN TRY
        IF dbo.fn_HYTPermiso(@UsuarioId,N'CONFIGURACION.CONSULTAR')=0 AND dbo.fn_HYTPermiso(@UsuarioId,N'PRESUPUESTOS.CONSULTAR')=0
            THROW 51300,N'[HYT_PERMISO] No tiene permiso para consultar las firmas.',1;
        SELECT COALESCE((SELECT firma.FirmaId,firma.Nombre,firma.ArchivoJson,firma.Predeterminada,firma.Activo
            FROM dbo.FirmaPresupuestoHYT AS firma ORDER BY firma.Predeterminada DESC,firma.Nombre FOR JSON PATH),N'[]') AS Datos;
    END TRY
    BEGIN CATCH
        DECLARE @ErrorMessage NVARCHAR(4000),@ErrorSeverity INT,@ErrorState INT;
        SELECT @ErrorMessage=ERROR_MESSAGE(),@ErrorSeverity=ERROR_SEVERITY(),@ErrorState=ERROR_STATE();
        RAISERROR(@ErrorMessage,@ErrorSeverity,@ErrorState);
    END CATCH
END;
GO
GRANT EXECUTE ON OBJECT::dbo.spu_FirmaPresupuestoHYT_Listar TO hyt_app_executor;
GO
