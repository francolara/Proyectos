-- =============================================
-- Author:        FRANCO LARA
-- Create date:   09/11/2026
-- Description:   spu_CargoTrabajadorHYT_Listar
-- =============================================
USE [SistemaWebHYT];
GO
CREATE OR ALTER PROCEDURE dbo.spu_CargoTrabajadorHYT_Listar
    @UsuarioId NVARCHAR(128)
AS
BEGIN
    SET NOCOUNT ON;
    BEGIN TRY
        IF dbo.fn_HYTPermiso(@UsuarioId, N'TRABAJADORES.CONSULTAR') = 0
            THROW 51000, N'No tiene permiso para consultar cargos.', 1;

        SELECT cargo.CargoId, cargo.Nombre, cargo.Activo, cargo.VersionFila
        FROM dbo.CargoTrabajadorHYT AS cargo
        ORDER BY cargo.Nombre, cargo.CargoId;
    END TRY
    BEGIN CATCH
        DECLARE @ErrorMessage NVARCHAR(4000), @ErrorSeverity INT, @ErrorState INT;
        SELECT @ErrorMessage = ERROR_MESSAGE(), @ErrorSeverity = ERROR_SEVERITY(), @ErrorState = ERROR_STATE();
        RAISERROR (@ErrorMessage, @ErrorSeverity, @ErrorState);
    END CATCH
END;
GO
