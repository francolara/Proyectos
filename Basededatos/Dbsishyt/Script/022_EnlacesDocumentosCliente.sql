-- =============================================
-- Author:        FRANCO LARA
-- Create date:   05/11/2026
-- Description:   Migra el enlace único a la lista una sola vez, sin borrar el dato anterior.
-- =============================================
USE [SistemaWebHYT];
GO
SET NOCOUNT ON;
SET XACT_ABORT ON;
BEGIN TRY
    BEGIN TRANSACTION;
    IF NOT EXISTS (SELECT 1 FROM sys.extended_properties AS propiedad WHERE propiedad.major_id=OBJECT_ID(N'dbo.ClienteEnlaceDocumento') AND propiedad.minor_id=0 AND propiedad.name=N'HYT_EnlaceAnteriorMigrado')
    BEGIN
        INSERT INTO dbo.ClienteEnlaceDocumento(ClienteId,Descripcion,Url,Orden,FechaRegistro,UsuarioRegistro)
        SELECT cliente.ClienteId,N'Documentos del cliente',cliente.UrlCargaDocumentos,0,cliente.FechaRegistro,cliente.UsuarioRegistro
        FROM dbo.Cliente AS cliente
        WHERE NULLIF(LTRIM(RTRIM(cliente.UrlCargaDocumentos)),N'') IS NOT NULL
          AND NOT EXISTS(SELECT 1 FROM dbo.ClienteEnlaceDocumento AS enlace WHERE enlace.ClienteId=cliente.ClienteId AND enlace.Url=cliente.UrlCargaDocumentos);
        EXEC sys.sp_addextendedproperty @name=N'HYT_EnlaceAnteriorMigrado',@value=1,@level0type=N'SCHEMA',@level0name=N'dbo',@level1type=N'TABLE',@level1name=N'ClienteEnlaceDocumento';
    END;
    COMMIT TRANSACTION;
END TRY
BEGIN CATCH
    IF XACT_STATE() <> 0 ROLLBACK TRANSACTION;
    DECLARE @ErrorMessage NVARCHAR(4000),@ErrorSeverity INT,@ErrorState INT;
    SELECT @ErrorMessage=ERROR_MESSAGE(),@ErrorSeverity=ERROR_SEVERITY(),@ErrorState=ERROR_STATE();
    RAISERROR(@ErrorMessage,@ErrorSeverity,@ErrorState);
END CATCH;
GO
