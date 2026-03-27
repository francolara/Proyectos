-- =============================================
-- Author:        FRANCO LARA
-- Create date:   26/03/2026
-- Description:   Procedimiento de auditoria central para operaciones CRUD.
-- =============================================

CREATE OR ALTER PROCEDURE dbo.Sp_Auditoria_Registrar
    @NegocioId INT = NULL,
    @Modulo NVARCHAR(50),
    @Accion NVARCHAR(20),
    @Entidad NVARCHAR(80),
    @EntidadId NVARCHAR(80),
    @Usuario NVARCHAR(200),
    @DetalleJson NVARCHAR(4000) = NULL
AS
BEGIN
    SET NOCOUNT ON;

    BEGIN TRY
        INSERT INTO dbo.BitacoraAuditoria
        (
            NegocioId,
            Modulo,
            Accion,
            Entidad,
            EntidadId,
            UsuarioId,
            UsuarioNombre,
            UsuarioCorreo,
            DetalleJson,
            FechaRegistro
        )
        VALUES
        (
            @NegocioId,
            @Modulo,
            @Accion,
            @Entidad,
            @EntidadId,
            @Usuario,
            @Usuario,
            NULL,
            @DetalleJson,
            SYSUTCDATETIME()
        );
    END TRY
    BEGIN CATCH
        DECLARE @ErrorMessage NVARCHAR(4000), @ErrorSeverity INT, @ErrorState INT;
        SELECT @ErrorMessage = ERROR_MESSAGE(), @ErrorSeverity = ERROR_SEVERITY(), @ErrorState = ERROR_STATE();
        RAISERROR (@ErrorMessage, @ErrorSeverity, @ErrorState);
    END CATCH
END
GO