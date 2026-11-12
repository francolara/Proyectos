-- Firma: FRANCO LARA | Create date: 10/11/2026 | Lee la configuración bancaria con el permiso de administración.
USE [SistemaWebHYT];
GO
CREATE OR ALTER PROCEDURE dbo.spu_TelecreditoHYT_Obtener
    @UsuarioId NVARCHAR(128)
AS
BEGIN
    SET NOCOUNT ON;
    BEGIN TRY
        IF dbo.fn_HYTPermiso(@UsuarioId,N'CONFIGURACION.CONSULTAR')=0
            THROW 51000,N'No tiene permiso para consultar Telecrédito.',1;
        SELECT (SELECT cuenta.TipoCuenta,cuenta.NumeroCuenta,cuenta.Moneda,cuenta.VersionFila
            FROM dbo.TelecreditoHYT AS cuenta WHERE cuenta.Id=1 FOR JSON PATH,WITHOUT_ARRAY_WRAPPER);
    END TRY
    BEGIN CATCH
        DECLARE @ErrorMessage NVARCHAR(4000), @ErrorSeverity INT, @ErrorState INT;
        SELECT @ErrorMessage=ERROR_MESSAGE(), @ErrorSeverity=ERROR_SEVERITY(), @ErrorState=ERROR_STATE();
        RAISERROR (@ErrorMessage,@ErrorSeverity,@ErrorState);
    END CATCH
END;
GO
