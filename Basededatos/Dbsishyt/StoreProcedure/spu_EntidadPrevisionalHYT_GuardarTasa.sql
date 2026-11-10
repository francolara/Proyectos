-- =============================================
-- Author:        FRANCO LARA
-- Create date:   10/11/2026
-- Description:   Actualiza el porcentaje de estimacion de una entidad previsional con auditoria.
-- =============================================
USE [SistemaWebHYT];
GO
CREATE OR ALTER PROCEDURE dbo.spu_EntidadPrevisionalHYT_GuardarTasa
    @UsuarioId NVARCHAR(128),
    @Entidad NVARCHAR(100),
    @Porcentaje DECIMAL(9,6)
AS
BEGIN
    SET NOCOUNT ON;
    SET XACT_ABORT ON;
    BEGIN TRY
        IF dbo.fn_HYTPermiso(@UsuarioId, N'OBLIGACIONES.CONSULTAR') = 0
           OR dbo.fn_HYTPermiso(@UsuarioId, N'OBLIGACIONES.EDITAR') = 0
            THROW 51000, N'No tiene permiso para guardar tasas previsionales.', 1;
        IF @Porcentaje IS NULL OR @Porcentaje <= 0 OR @Porcentaje > 100 OR @Porcentaje <> ROUND(@Porcentaje, 4)
            THROW 51000, N'Ingrese un porcentaje valido.', 1;

        BEGIN TRANSACTION;
        UPDATE dbo.EntidadPrevisionalHYT
        SET PorcentajeEstimado = @Porcentaje
        WHERE Nombre = @Entidad;
        IF @@ROWCOUNT = 0
            THROW 51000, N'La entidad previsional no existe.', 1;

        DECLARE @OperacionId UNIQUEIDENTIFIER = NEWID();
        DECLARE @Cambios NVARCHAR(MAX) = (SELECT @Porcentaje AS PorcentajeEstimado FOR JSON PATH, WITHOUT_ARRAY_WRAPPER);
        EXEC dbo.spu_RegistrarAuditoria @UsuarioId, @OperacionId, N'EntidadPrevisionalHYT', @Entidad,
             N'MODIFICAR', @CambiosJson = @Cambios;
        COMMIT TRANSACTION;
        SELECT CONVERT(BIGINT, 1) AS Guardado;
    END TRY
    BEGIN CATCH
        IF XACT_STATE() <> 0 ROLLBACK TRANSACTION;
        DECLARE @ErrorMessage NVARCHAR(4000), @ErrorSeverity INT, @ErrorState INT;
        SELECT @ErrorMessage = ERROR_MESSAGE(), @ErrorSeverity = ERROR_SEVERITY(), @ErrorState = ERROR_STATE();
        RAISERROR (@ErrorMessage, @ErrorSeverity, @ErrorState);
    END CATCH
END;
GO
