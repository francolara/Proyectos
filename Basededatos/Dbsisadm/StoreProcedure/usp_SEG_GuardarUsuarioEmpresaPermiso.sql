-- =============================================
-- Author:        FRANCO LARA / Codex
-- Create date:   10/07/2026
-- Description:   Guarda o actualiza el override de permisos del usuario para un modulo de alcance empresa.
-- =============================================

CREATE OR ALTER PROCEDURE dbo.usp_SEG_GuardarUsuarioEmpresaPermiso
    @IdUsuarioEmpresa INT,
    @IdModuloSistema INT,
    @PuedeVer BIT = NULL,
    @PuedeCrear BIT = NULL,
    @PuedeEditar BIT = NULL,
    @PuedeEliminar BIT = NULL,
    @UsuarioRegistro NVARCHAR(450) = NULL
AS
BEGIN

    SET NOCOUNT ON;

    BEGIN TRY

        IF NOT EXISTS
        (
            SELECT 1
            FROM dbo.SEG_ModuloSistema AS ms
            WHERE ms.IdModuloSistema = @IdModuloSistema
              AND ms.AlcanceModulo = 'EMPRESA'
              AND ms.Estado = 1
        )
        BEGIN
            RAISERROR (N'El modulo no existe, no esta activo o no pertenece al alcance EMPRESA.', 16, 1);
            RETURN;
        END;

        IF EXISTS
        (
            SELECT 1
            FROM dbo.SEG_UsuarioEmpresaPermiso AS uep
            WHERE uep.IdUsuarioEmpresa = @IdUsuarioEmpresa
              AND uep.IdModuloSistema = @IdModuloSistema
        )
        BEGIN
            UPDATE dbo.SEG_UsuarioEmpresaPermiso
            SET PuedeVer = @PuedeVer,
                PuedeCrear = @PuedeCrear,
                PuedeEditar = @PuedeEditar,
                PuedeEliminar = @PuedeEliminar,
                Estado = 1,
                UsuarioRegistro = @UsuarioRegistro
            WHERE IdUsuarioEmpresa = @IdUsuarioEmpresa
              AND IdModuloSistema = @IdModuloSistema;
        END
        ELSE
        BEGIN
            INSERT INTO dbo.SEG_UsuarioEmpresaPermiso
            (
                IdUsuarioEmpresa,
                IdModuloSistema,
                PuedeVer,
                PuedeCrear,
                PuedeEditar,
                PuedeEliminar,
                Estado,
                UsuarioRegistro
            )
            VALUES
            (
                @IdUsuarioEmpresa,
                @IdModuloSistema,
                @PuedeVer,
                @PuedeCrear,
                @PuedeEditar,
                @PuedeEliminar,
                1,
                @UsuarioRegistro
            );
        END;

    END TRY

    BEGIN CATCH

        DECLARE @ErrorMessage NVARCHAR(4000)
        DECLARE @ErrorSeverity INT
        DECLARE @ErrorState INT

        SELECT
            @ErrorMessage = ERROR_MESSAGE(),
            @ErrorSeverity = ERROR_SEVERITY(),
            @ErrorState = ERROR_STATE()

        RAISERROR (@ErrorMessage, @ErrorSeverity, @ErrorState)

    END CATCH

END
