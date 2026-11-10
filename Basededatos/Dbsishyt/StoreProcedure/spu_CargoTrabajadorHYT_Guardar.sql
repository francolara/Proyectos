-- =============================================
-- Author:        FRANCO LARA
-- Create date:   09/11/2026
-- Description:   spu_CargoTrabajadorHYT_Guardar
-- =============================================
USE [SistemaWebHYT];
GO
CREATE OR ALTER PROCEDURE dbo.spu_CargoTrabajadorHYT_Guardar
    @UsuarioId NVARCHAR(128),
    @Nombre NVARCHAR(100),
    @Activo BIT,
    @CargoId BIGINT = NULL,
    @VersionFila BINARY(8) = NULL
AS
BEGIN
    SET NOCOUNT ON;
    SET XACT_ABORT ON;
    BEGIN TRY
        IF dbo.fn_HYTPermiso(@UsuarioId, N'TRABAJADORES.EDITAR') = 0
            THROW 51000, N'No tiene permiso para guardar cargos.', 1;

        SET @Nombre = NULLIF(TRIM(@Nombre), N'');
        IF @Nombre IS NULL OR @Activo IS NULL
            THROW 51000, N'Indique el nombre y el estado del cargo.', 1;

        BEGIN TRANSACTION;
        DECLARE @Accion NVARCHAR(60) = N'MODIFICAR';
        IF @CargoId IS NULL
        BEGIN
            INSERT INTO dbo.CargoTrabajadorHYT (Nombre, Activo) VALUES (@Nombre, @Activo);
            SET @CargoId = SCOPE_IDENTITY();
            SET @Accion = N'CREAR';
        END
        ELSE
        BEGIN
            UPDATE dbo.CargoTrabajadorHYT SET Nombre = @Nombre, Activo = @Activo
            WHERE CargoId = @CargoId AND VersionFila = @VersionFila;
            IF @@ROWCOUNT = 0
                THROW 51000, N'El cargo cambio o no existe.', 1;
        END;

        DECLARE @CorrelacionId UNIQUEIDENTIFIER = NEWID();
        DECLARE @RegistroId NVARCHAR(100) = CONVERT(NVARCHAR(100), @CargoId);
        DECLARE @Cambios NVARCHAR(MAX) = (SELECT @Nombre AS Nombre, @Activo AS Activo FOR JSON PATH, WITHOUT_ARRAY_WRAPPER);
        EXEC dbo.spu_RegistrarAuditoria @UsuarioId, @CorrelacionId, N'CargoTrabajadorHYT',
             @RegistroId, @Accion, @CambiosJson = @Cambios;
        COMMIT TRANSACTION;
        SELECT @CargoId AS CargoId;
    END TRY
    BEGIN CATCH
        IF XACT_STATE() <> 0 ROLLBACK TRANSACTION;
        DECLARE @ErrorMessage NVARCHAR(4000), @ErrorSeverity INT, @ErrorState INT;
        SELECT @ErrorMessage = ERROR_MESSAGE(), @ErrorSeverity = ERROR_SEVERITY(), @ErrorState = ERROR_STATE();
        RAISERROR (@ErrorMessage, @ErrorSeverity, @ErrorState);
    END CATCH
END;
GO
