-- Firma: FRANCO LARA | Create date: 14/11/2026 | Aplica las transiciones de factoring en orden y las registra en auditoría.
USE [SistemaWebHYT];
GO
CREATE OR ALTER PROCEDURE dbo.spu_ComprobanteHYT_FactoringCambiarEstado
    @UsuarioId NVARCHAR(128), @Id BIGINT, @Estado NVARCHAR(12)
AS
BEGIN
    SET NOCOUNT ON;
    BEGIN TRY
        SET XACT_ABORT ON;
        IF dbo.fn_HYTPermiso(@UsuarioId,N'FACTORING.CONSULTAR')=0 OR dbo.fn_HYTPermiso(@UsuarioId,N'FACTORING.EDITAR')=0
            THROW 51600,N'[HYT_PERMISO] No tiene permiso para cambiar el estado del lote.',1;
        IF @Estado IS NULL OR @Estado NOT IN (N'ACEPTADO',N'CERRADO')
            THROW 51601,N'[HYT_VALIDACION] Estado de factoring no válido.',1;

        DECLARE @Anterior NVARCHAR(12), @Operacion UNIQUEIDENTIFIER=NEWID(), @Registro NVARCHAR(50)=CONVERT(NVARCHAR(50),@Id);
        BEGIN TRANSACTION;
        SELECT @Anterior=lote.Estado FROM dbo.FactoringLoteHYT AS lote WITH(UPDLOCK,HOLDLOCK) WHERE lote.LoteId=@Id;
        IF @Anterior IS NULL THROW 51601,N'[HYT_VALIDACION] No se encontró el lote de factoring.',1;
        IF @Anterior<>@Estado
        BEGIN
            IF NOT ((@Anterior=N'REGISTRADO' AND @Estado=N'ACEPTADO') OR (@Anterior=N'ACEPTADO' AND @Estado=N'CERRADO'))
                THROW 51601,N'[HYT_VALIDACION] El lote debe pasar de Registrado a Aceptado y luego a Cerrado.',1;
            UPDATE dbo.FactoringLoteHYT SET Estado=@Estado WHERE LoteId=@Id;
            DECLARE @Accion NVARCHAR(20)=CASE WHEN @Estado=N'ACEPTADO' THEN N'ACEPTAR' ELSE N'CERRAR' END;
            EXEC dbo.spu_RegistrarAuditoria @UsuarioId,@Operacion,N'FactoringLoteHYT',@Registro,@Accion,@Anterior,@Estado;
        END;
        COMMIT TRANSACTION;
        SELECT @Id;
    END TRY
    BEGIN CATCH
        IF XACT_STATE()<>0 ROLLBACK TRANSACTION;
        DECLARE @ErrorMessage NVARCHAR(4000),@ErrorSeverity INT,@ErrorState INT;
        SELECT @ErrorMessage=ERROR_MESSAGE(),@ErrorSeverity=ERROR_SEVERITY(),@ErrorState=ERROR_STATE();
        RAISERROR (@ErrorMessage,@ErrorSeverity,@ErrorState);
    END CATCH
END;
GO
GRANT EXECUTE ON OBJECT::dbo.spu_ComprobanteHYT_FactoringCambiarEstado TO hyt_app_executor;
GO
