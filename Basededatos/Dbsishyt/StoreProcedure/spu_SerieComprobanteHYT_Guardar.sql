-- Firma: FRANCO LARA | Create date: 17/11/2026 | Agrega una serie por tipo y registra auditoría sin duplicados.
USE [SistemaWebHYT];
GO
CREATE OR ALTER PROCEDURE dbo.spu_SerieComprobanteHYT_Guardar
    @UsuarioId NVARCHAR(128),@Tipo VARCHAR(20),@Serie NVARCHAR(100)
AS
BEGIN
    SET NOCOUNT ON;
    SET XACT_ABORT ON;
    BEGIN TRY
        IF dbo.fn_HYTPermiso(@UsuarioId,N'CONFIGURACION.CONSULTAR')=0 OR dbo.fn_HYTPermiso(@UsuarioId,N'CONFIGURACION.EDITAR')=0
            THROW 51600,N'[HYT_PERMISO] No tiene permiso para editar la configuración.',1;
        SET @Serie=UPPER(TRIM(@Serie));
        IF @Tipo IS NULL OR @Tipo NOT IN ('01','03','07') OR @Serie IS NULL OR LEN(@Serie) NOT BETWEEN 1 AND 30
            OR @Serie COLLATE Latin1_General_100_BIN2 LIKE N'%[^A-Z0-9]%'
            THROW 51601,N'[HYT_VALIDACION] Selecciona un tipo y escribe una serie alfanumérica de hasta 30 caracteres.',1;
        BEGIN TRANSACTION;
        IF EXISTS(SELECT 1 FROM dbo.SerieComprobanteHYT WITH(UPDLOCK,HOLDLOCK) WHERE Tipo=@Tipo AND Serie=@Serie)
            THROW 51601,N'[HYT_VALIDACION] Esa serie ya existe para el tipo seleccionado.',1;
        INSERT INTO dbo.SerieComprobanteHYT(Tipo,Serie,UsuarioRegistro) VALUES(@Tipo,@Serie,@UsuarioId);
        DECLARE @Registro NVARCHAR(100)=CONCAT(@Tipo,N'-',@Serie),@Operacion UNIQUEIDENTIFIER=NEWID();
        EXEC dbo.spu_RegistrarAuditoria @UsuarioId,@Operacion,N'SerieComprobanteHYT',@Registro,N'CREAR';
        COMMIT TRANSACTION;
    END TRY
    BEGIN CATCH
        IF XACT_STATE()<>0 ROLLBACK TRANSACTION;
        DECLARE @ErrorMessage NVARCHAR(4000),@ErrorSeverity INT,@ErrorState INT;
        SELECT @ErrorMessage=ERROR_MESSAGE(),@ErrorSeverity=ERROR_SEVERITY(),@ErrorState=ERROR_STATE();
        RAISERROR (@ErrorMessage,@ErrorSeverity,@ErrorState);
    END CATCH
END;
GO
GRANT EXECUTE ON OBJECT::dbo.spu_SerieComprobanteHYT_Guardar TO hyt_app_executor;
GO
